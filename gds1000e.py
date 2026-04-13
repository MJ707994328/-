from __future__ import annotations

import glob
import os
import select
import termios
import time
from dataclasses import dataclass
from pathlib import Path

from PIL import Image


KNOWN_DISPLAY_SIZES = {
    800 * 480: (800, 480),
    640 * 480: (640, 480),
    320 * 240: (320, 240),
}


@dataclass(slots=True)
class ScopeIdentity:
    port: str
    manufacturer: str
    model: str
    serial_number: str
    firmware: str
    raw: str


class GDS1000ESerialClient:
    """Minimal SCPI client for GDS-1000E scopes over USB CDC/ttyACM."""

    def __init__(self, port: str, baudrate: int = 115200, poll_interval: float = 0.1) -> None:
        self.port = port
        self.baudrate = baudrate
        self.poll_interval = poll_interval
        self.fd: int | None = None

    def __enter__(self) -> "GDS1000ESerialClient":
        self.open()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def open(self) -> None:
        if self.fd is not None:
            return

        baud_constant = self._baudrate_to_termios(self.baudrate)
        fd = os.open(self.port, os.O_RDWR | os.O_NOCTTY | os.O_NONBLOCK)
        attrs = termios.tcgetattr(fd)
        attrs[0] = 0
        attrs[1] = 0
        attrs[2] = termios.CLOCAL | termios.CREAD | termios.CS8
        attrs[3] = 0
        attrs[4] = baud_constant
        attrs[5] = baud_constant
        attrs[6][termios.VMIN] = 0
        attrs[6][termios.VTIME] = 1
        termios.tcsetattr(fd, termios.TCSANOW, attrs)
        termios.tcflush(fd, termios.TCIOFLUSH)
        self.fd = fd

    def close(self) -> None:
        if self.fd is None:
            return
        os.close(self.fd)
        self.fd = None

    @staticmethod
    def _baudrate_to_termios(baudrate: int) -> int:
        name = f"B{baudrate}"
        if not hasattr(termios, name):
            raise ValueError(f"Unsupported baudrate: {baudrate}")
        return getattr(termios, name)

    def _require_open(self) -> int:
        if self.fd is None:
            raise RuntimeError("Serial port is not open.")
        return self.fd

    def _drain_until_idle(self, deadline_sec: float = 3.0, idle_rounds_limit: int = 5) -> bytes:
        fd = self._require_open()
        buffer = bytearray()
        idle_rounds = 0
        deadline = time.time() + deadline_sec

        while time.time() < deadline and idle_rounds < idle_rounds_limit:
            readable, _, _ = select.select([fd], [], [], self.poll_interval)
            if not readable:
                idle_rounds += 1
                continue

            chunk = os.read(fd, 65536)
            if chunk:
                buffer.extend(chunk)
                idle_rounds = 0
            else:
                idle_rounds += 1

        return bytes(buffer)

    def _write_command(self, command: str) -> None:
        fd = self._require_open()
        os.write(fd, command.encode("ascii") + b"\n")

    def query_text(self, command: str, timeout: float = 2.0) -> str:
        self._drain_until_idle(deadline_sec=0.5)
        self._write_command(command)
        time.sleep(0.05)
        raw = self._drain_until_idle(deadline_sec=timeout)
        if not raw:
            raise TimeoutError(f"No response for {command!r} on {self.port}")
        return raw.decode("utf-8", errors="replace").strip()

    def query_binary_block(self, command: str, timeout: float = 8.0) -> bytes:
        self._drain_until_idle(deadline_sec=1.0)
        self._write_command(command)
        time.sleep(0.1)
        raw = self._drain_until_idle(deadline_sec=timeout, idle_rounds_limit=8)
        if not raw:
            raise TimeoutError(f"No binary response for {command!r} on {self.port}")

        start = raw.find(b"#")
        if start < 0:
            raise ValueError(f"Binary block header not found in response from {self.port}")

        raw = raw[start:]
        if len(raw) < 2:
            raise ValueError("Incomplete binary block header.")

        digits = int(chr(raw[1]))
        header_end = 2 + digits
        if len(raw) < header_end:
            raise ValueError("Incomplete binary block length field.")

        expected_length = int(raw[2:header_end].decode("ascii"))
        data_end = header_end + expected_length
        if len(raw) < data_end:
            raise ValueError(
                f"Incomplete binary payload: expected {expected_length} bytes, got {len(raw) - header_end}"
            )

        return raw[header_end:data_end]

    def identify(self) -> ScopeIdentity:
        raw = self.query_text("*IDN?")
        parts = [part.strip() for part in raw.split(",")]
        while len(parts) < 4:
            parts.append("")
        manufacturer, model, serial_number, firmware = parts[:4]
        return ScopeIdentity(
            port=self.port,
            manufacturer=manufacturer,
            model=model,
            serial_number=serial_number,
            firmware=firmware,
            raw=raw,
        )

    def capture_display_rle(self) -> bytes:
        return self.query_binary_block(":DISPlay:OUTPut?")

    def capture_display_image(self) -> Image.Image:
        rle_data = self.capture_display_rle()
        rgb_bytes, width, height = decode_display_output(rle_data)
        return Image.frombytes("RGB", (width, height), rgb_bytes)

    def save_display_image(self, output_path: str | Path) -> Path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        image = self.capture_display_image()
        image.save(output_path)
        return output_path


def list_candidate_ports() -> list[str]:
    candidates = sorted(
        set(glob.glob("/dev/ttyACM*") + glob.glob("/dev/ttyUSB*"))
    )
    return candidates


def discover_scopes() -> list[ScopeIdentity]:
    scopes: list[ScopeIdentity] = []
    for port in list_candidate_ports():
        try:
            with GDS1000ESerialClient(port) as scope:
                identity = scope.identify()
        except Exception:
            continue

        normalized = f"{identity.manufacturer},{identity.model}".upper()
        if "GW" in normalized and "GDS" in normalized:
            scopes.append(identity)

    return scopes


def autodetect_scope() -> ScopeIdentity:
    scopes = discover_scopes()
    if not scopes:
        raise RuntimeError("No GW Instek GDS oscilloscope found on /dev/ttyACM* or /dev/ttyUSB*")
    return scopes[0]


def decode_display_output(rle_data: bytes) -> tuple[bytes, int, int]:
    if len(rle_data) % 4 != 0:
        raise ValueError(f"Invalid RLE payload length: {len(rle_data)}")

    total_pixels = 0
    runs: list[tuple[int, int]] = []
    for index in range(0, len(rle_data), 4):
        count = int.from_bytes(rle_data[index:index + 2], "little")
        color = int.from_bytes(rle_data[index + 2:index + 4], "little")
        runs.append((count, color))
        total_pixels += count

    if total_pixels not in KNOWN_DISPLAY_SIZES:
        raise ValueError(f"Unsupported display size, decoded pixel count = {total_pixels}")

    width, height = KNOWN_DISPLAY_SIZES[total_pixels]
    rgb_bytes = bytearray(total_pixels * 3)
    offset = 0

    for count, color in runs:
        red = ((color >> 11) & 0x1F) * 255 // 31
        green = ((color >> 5) & 0x3F) * 255 // 63
        blue = (color & 0x1F) * 255 // 31
        triplet = bytes((red, green, blue))
        rgb_bytes[offset:offset + (count * 3)] = triplet * count
        offset += count * 3

    return bytes(rgb_bytes), width, height
