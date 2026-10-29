# HKUST slide templates (for CourseProj lectures)

## Downloaded locally

| File | Source | Notes |
|------|--------|--------|
| `HKUST_Teaching.pptx` | [GECO Brand Materials – Teaching](https://geco.hkust.edu.hk/brand-materials) | Official corporate **Teaching** theme (~15 MB, verified). Open in PowerPoint → **Save As** → `.potx` if you want a reusable template. |

MTPC Template 16 (`.potx`, ~56 MB) did not finish downloading here; use the link below in a browser if you want that variant.

## Official download pages (more variants)

- **Brand Assets Unit – PowerPoint templates:** https://brand.hkust.edu.hk/ (see “Brand PowerPoint Templates” and [Brand Guidelines PDF](https://brand.hkust.edu.hk/sites/default/files/bm_folder/HKUST_Brand_guideline.pdf) — 16:9, Muli or Arial)
- **GECO – themes:** Architecture, Seaview, Sports, Sundial, **Teaching**, school packs (SENG, DSCI, …): https://geco.hkust.edu.hk/brand-materials
- **MTPC – `.potx` numbered templates (e.g. Template 16):** https://mtpc.hkust.edu.hk/node/297

Mirror (if GECO is slow): `https://pao.ust.hk/Bdownload/HKUSTpowerpoint/HKUST_PPT_Corporate/Teaching.pptx`

## Not HKUST (do not use as official template)

- `~/Desktop/Research_HKUST.pptx` — metadata says **Columbia Engineering** template, not HKUST.

Dark full-bleed layouts (section dividers, and the second content set — template slides 11–15) are not used for teaching slides. Cover and the Thank you / Q&A ending may stay dark.

## Minimum slides for `build_rl_section.py` (after you trim the template)

You may delete unused demo pages. The builder skips dark full-bleed slides and picks light prototypes. Keep at least:

1. **Cover** (first slide; dark background is fine)
2. **One light “Title and Content”** example (the earlier content set, not slides 11–15)
3. **One light “Two Content”** example with two content placeholders

Optional: four-card slide, **Thank you**, **Q&A**. Section pages are a light Section Header, not the dark divider styles.

Regenerate lecture slides: `python3 slides/build_rl_section.py` → `slides/CourseProject.pptx`.
