from __future__ import annotations

from datetime import datetime
from pathlib import Path

from gds1000e import GDS1000ESerialClient, autodetect_scope


def main() -> None:
    identity = autodetect_scope()
    output_dir = Path(__file__).with_name("captures")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"{identity.model}_{timestamp}.png"

    with GDS1000ESerialClient(identity.port) as scope:
        saved_path = scope.save_display_image(output_path)

    print(f"设备: {identity.raw}")
    print(f"截图已保存到: {saved_path}")


if __name__ == "__main__":
    main()
