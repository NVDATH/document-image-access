"""Write the notes of one version to a file, for the GitHub release body.

Usage: python scripts/extract_release_notes.py VERSION release_notes.md
Takes the body of the "## VERSION" section of changelog.md (up to the next "## " or the end of the file).
Falls back to a one-line note when the section is missing.
"""
import re
import sys


def notes_for(text, version):
    parts = re.split(r"(?m)^## +(.*)$", text)
    for i in range(1, len(parts) - 1, 2):
        if re.match(rf"v?{re.escape(version)}(?:[\s(\[-]|$)", parts[i].strip()):
            return parts[i + 1].strip("\n")
    return ""


def main():
    version, out = sys.argv[1], sys.argv[2]
    try:
        with open("changelog.md", encoding="utf-8") as f:
            notes = notes_for(f.read(), version)
    except FileNotFoundError:
        notes = ""
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(notes.strip() + "\n" if notes.strip() else f"Release for version {version}\n")


if __name__ == "__main__":
    main()
