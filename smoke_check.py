from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parent

REQUIRED_FILES = [
    ROOT / "_version.py",
    ROOT / "generate_version_info.py",
    ROOT / "file_version_info.txt",
    ROOT / "BingWallpaper.spec",
    ROOT / "InstallBingWallpaper.spec",
    ROOT / "build_installer.ps1",
]


def fail(message: str) -> int:
    print(f"ERROR: {message}")
    return 1


def parse_version() -> str:
    text = (ROOT / "_version.py").read_text(encoding="utf-8")
    match = re.search(r"__version__\s*=\s*['\"]([^'\"]+)['\"]", text)
    if not match:
        raise ValueError("Unable to parse __version__ from _version.py")

    raw = match.group(1).strip().lstrip("v")
    parts = []
    for chunk in raw.split("."):
        digits = "".join(ch for ch in chunk if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 4:
        parts.append(0)
    return ".".join(str(p) for p in parts[:4])


def main() -> int:
    for file_path in REQUIRED_FILES:
        if not file_path.exists():
            return fail(f"Required file missing: {file_path.name}")

    try:
        version = parse_version()
    except Exception as exc:
        return fail(str(exc))

    version_info_text = (ROOT / "file_version_info.txt").read_text(encoding="utf-8")
    if f"StringStruct('FileVersion', '{version}')" not in version_info_text:
        return fail("file_version_info.txt FileVersion is out of sync")
    if f"StringStruct('ProductVersion', '{version}')" not in version_info_text:
        return fail("file_version_info.txt ProductVersion is out of sync")

    app_spec = (ROOT / "BingWallpaper.spec").read_text(encoding="utf-8")
    installer_spec = (ROOT / "InstallBingWallpaper.spec").read_text(encoding="utf-8")
    if "version='file_version_info.txt'" not in app_spec:
        return fail("BingWallpaper.spec is missing version resource wiring")
    if "version='file_version_info.txt'" not in installer_spec:
        return fail("InstallBingWallpaper.spec is missing version resource wiring")

    build_script = (ROOT / "build_installer.ps1").read_text(encoding="utf-8")
    if "generate_version_info.py" not in build_script:
        return fail("build_installer.ps1 does not generate version metadata")

    print("OK: smoke checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
