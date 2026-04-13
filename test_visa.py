from __future__ import annotations

try:
    import pyvisa
except Exception:
    pyvisa = None

from gds1000e import discover_scopes, list_candidate_ports


def main() -> None:
    print("=== PyVISA 扫描 ===")
    if pyvisa is None:
        print("pyvisa 未安装，跳过 VISA 扫描。")
    else:
        try:
            rm = pyvisa.ResourceManager("@py")
            resources = rm.list_resources()
            print(f"PyVISA 发现的资源: {resources}")
        except Exception as exc:
            print(f"PyVISA 扫描失败: {exc}")

    print("\n=== 串口扫描 ===")
    ports = list_candidate_ports()
    print(f"候选串口: {ports}")

    scopes = discover_scopes()
    if not scopes:
        print("没有在 /dev/ttyACM* 或 /dev/ttyUSB* 上发现 GW Instek GDS 示波器。")
        return

    for identity in scopes:
        print(f"{identity.port} -> {identity.raw}")


if __name__ == "__main__":
    main()
