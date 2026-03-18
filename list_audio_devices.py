"""List available DirectShow audio input devices for ffmpeg."""

from __future__ import annotations

import re
import subprocess


def decode_ffmpeg_output(raw_bytes: bytes) -> str:
    for encoding in ("utf-8", "gbk", "cp936"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("latin1", errors="ignore")


def list_audio_devices() -> list[str]:
    cmd = ["ffmpeg", "-hide_banner", "-list_devices", "true", "-f", "dshow", "-i", "dummy"]
    result = subprocess.run(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        check=False,
    )

    output = decode_ffmpeg_output(result.stderr)
    devices: list[str] = []
    in_audio_section = False
    saw_audio_section = False

    for line in output.splitlines():
        if "DirectShow audio devices" in line:
            in_audio_section = True
            saw_audio_section = True
            continue
        if in_audio_section and "DirectShow video devices" in line:
            break

        if not in_audio_section:
            continue
        if "Alternative name" in line:
            continue

        match = re.search(r'"(.+?)"', line)
        if match:
            name = match.group(1).strip()
            if name not in devices:
                devices.append(name)

    if not devices and not saw_audio_section:
        for line in output.splitlines():
            if "dshow" not in line.lower():
                continue
            if "Alternative name" in line:
                continue
            match = re.search(r'"(.+?)"', line)
            if match:
                name = match.group(1).strip()
                if name not in devices:
                    devices.append(name)

    return devices


def main() -> None:
    try:
        devices = list_audio_devices()
    except FileNotFoundError:
        print("ffmpeg not found. Please install ffmpeg and add it to PATH.")
        return
    except Exception as exc:
        print(f"Failed to query audio devices: {exc}")
        return

    print("=" * 60)
    print("Available audio devices")
    print("=" * 60)
    if not devices:
        print("No DirectShow audio devices were detected.")
    else:
        for idx, name in enumerate(devices, start=1):
            print(f"{idx:>2}. {name}")

    print("\nUse one of the following in recorder_config.json:")
    if devices:
        for name in devices:
            print(f'  "audio_device": "audio={name}"')
    else:
        print('  "audio_device": "audio=Your Device Name"')


if __name__ == "__main__":
    main()
