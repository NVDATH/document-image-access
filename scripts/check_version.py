"""Fail when the version being released differs from addon_version in buildVars.py.

Usage: python scripts/check_version.py X.Y.Z   (the tag without the leading v)
The Add-on Store reads the version from the manifest, so a tag that disagrees with it would publish a
build under the wrong version.
"""
import os
import sys

sys.path.insert(0, os.getcwd())  # run from the repository root; sys.path[0] would be scripts/
import buildVars  # noqa: E402

expected = sys.argv[1]
actual = buildVars.addon_info["addon_version"]
if expected != actual:
    sys.exit(f"The tag says {expected} but buildVars.py says {actual}. Fix one of them and tag again.")
print(f"version {actual} matches the tag")
