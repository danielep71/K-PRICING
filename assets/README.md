# 🎨 Assets

`assets/` is the canonical home for versioned, non-code project assets.

Typical contents include:

- social-preview artwork;
- diagrams and screenshots used by repository documentation;
- icons or Ribbon resources that are source inputs; and
- editable source artwork required to reproduce published visuals.

<!-- template:remove:start -->
The template's public presentation uses
[`social-preview.png`](social-preview.png), rendered at GitHub's recommended
1280×640 resolution from the editable
[`social-preview.svg`](social-preview.svg). Both files are template-only and are
removed during initialization unless `SOCIAL_PREVIEW_PATH` explicitly selects
the PNG. Generated projects should replace the artwork with their own identity
before retaining it.
<!-- template:remove:end -->

Keep source assets reviewable, licensed, and free of personal or confidential information. Use meaningful filenames and update every reference when an asset moves.

An `images/` directory is legitimate only when an established documentation renderer, package layout, or stable public URL requires that name. Choose `assets/` or `images/` for the same role—never both.

Generated release packages do not belong here. Use an ignored local `dist/` directory or workflow artifacts, then publish certified binaries through GitHub Releases. Track `dist/` only when a project's explicit release contract treats generated distributions as versioned source and verifies their provenance.

Delete this README only if real assets and equivalent documentation make the directory's role equally explicit.
