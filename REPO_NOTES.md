# Setting up the repository

Create the GitHub repository `NVDATH/document-image-access` (every URL in this package assumes that name),
then copy the contents of this package into it.

These come unchanged from the NVDA add-on template you already use, so they are not repeated here:

- `site_scons/site_tools/gettexttool/` (used by `SConstruct`)
- `manifest.ini.tpl` and `manifest-translated.ini.tpl`
- `style.css`
- `LICENSE`

Notes:

- The add-on ID is `docImages` (`addon_name` in `buildVars.py`). It cannot change after the first release.
- The first release is tagged `v0.6.2`. The tag must equal `addon_version` in `buildVars.py`.
- `.github/workflows/ci.yaml` and `notify_translators.yml` run by hand (Actions tab, Run workflow).
  `build_addon.yaml` runs when a tag is pushed.
