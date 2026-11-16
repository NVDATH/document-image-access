"""Copy one section of changelog.md into buildVars.py, between the CHANGELOG-START/END markers.

Which section: the tag being built (GITHUB_REF=refs/tags/vX.Y.Z -> "## X.Y.Z"), otherwise the topmost
"## " section. Run from the repository root: python scripts/sync_changelog.py
"""
import os
import re
import sys

CHANGELOG = "changelog.md"
BUILD_VARS = "buildVars.py"
MARKERS = re.compile(r"(?P<start>[ \t]*# CHANGELOG-START\n)(?P<body>.*?)(?P<end>[ \t]*# CHANGELOG-END)", re.S)


def sections(text):
    """[(heading, body)] for every "## " section."""
    parts = re.split(r"(?m)^## +(.*)$", text)
    return [(parts[i].strip(), parts[i + 1].strip("\n")) for i in range(1, len(parts) - 1, 2)]


def wanted_version():
    ref = os.environ.get("GITHUB_REF", "")
    return ref[len("refs/tags/v"):] if ref.startswith("refs/tags/v") else None


def pick(found, version):
    if version:
        for heading, body in found:
            if re.match(rf"v?{re.escape(version)}(?:[\s(\[-]|$)", heading):
                return heading, body
        sys.exit(f"changelog.md has no '## {version}' section")
    if not found:
        sys.exit("changelog.md has no '## ' section")
    return found[0]


def python_string(body):
    body = body.replace("\\", "\\\\").replace('"""', '\\"\\"\\"')
    return f'        """{body}"""\n'


def main():
    with open(CHANGELOG, encoding="utf-8") as f:
        heading, body = pick(sections(f.read()), wanted_version())
    with open(BUILD_VARS, encoding="utf-8") as f:
        source = f.read()
    if not MARKERS.search(source):
        sys.exit("CHANGELOG-START/END markers not found in buildVars.py")
    updated = MARKERS.sub(lambda m: m.group("start") + python_string(body) + m.group("end"), source, count=1)
    if updated != source:
        with open(BUILD_VARS, "w", encoding="utf-8", newline="\n") as f:
            f.write(updated)
    print(f"changelog section used: {heading}")


if __name__ == "__main__":
    main()
