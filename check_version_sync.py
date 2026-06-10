from pathlib import Path
import sys

import generate_version_info

ROOT = Path(__file__).resolve().parent
VERSION_FILE = ROOT / "_version.py"
VERSION_INFO_FILE = ROOT / "file_version_info.txt"


def main():
    if not VERSION_FILE.exists():
        print("ERROR: _version.py not found.")
        return 1

    before = VERSION_INFO_FILE.read_text(encoding="utf-8") if VERSION_INFO_FILE.exists() else ""
    generate_version_info.main()
    after = VERSION_INFO_FILE.read_text(encoding="utf-8") if VERSION_INFO_FILE.exists() else ""

    if before != after:
        print("ERROR: file_version_info.txt was out of sync with _version.py.")
        print("It has been regenerated. Please stage file_version_info.txt and commit again.")
        return 1

    print("OK: version metadata is in sync.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
