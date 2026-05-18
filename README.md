# WCG Panel Takeoff Tool v2

Extracts panel schedules from SolidWorks STEP files **or WCG supplier-drawing PDFs** and exports to the Smart Space supplier pricing template.

## Quick Start

### First-time setup (one-off)

1. Install Python 3.10+ from https://www.python.org/downloads/ (tick "Add to PATH")
2. Open Command Prompt in this folder and run:
   ```
   pip install -r requirements.txt
   ```
   Note: `cadquery` is ~200MB as it includes the OpenCASCADE geometry kernel.
3. Double-click **Launch Panel Takeoff.bat**

### Daily use

1. Double-click **Launch Panel Takeoff.bat** — opens in browser at http://localhost:8501
2. Set project details and material assignments in the left sidebar
3. **STEP tab** — upload STEP files (one per assembly/drawing) and click **Extract Panels**
   **PDF tab** — upload WCG supplier-drawing PDFs (with `Component:` description blocks) and click **Extract Panels from PDFs**
4. Review the editable table — check auto-notes for items needing attention
5. Click **Generate Smart Space XLSX** and download

## PDF Takeoff (WCG supplier-drawing format)

The PDF parser reads supplier drawings that follow the WCG template — each
component has a `Component:` description block with **Material**, **Thickness**,
**Edge Finish**, and **Qty** fields, plus dimension labels in the drawing.

- Component name → part type (mapped via keywords: Worktop, Pelmet, Plinth,
  Back, LH/RH Panel, End, Top, Base, Shelf, Door…).
- **Edge Finish** → edging code: "Edge Tape All Around" → `LLWW`,
  "Edge Tape Where Indicated" → `L`, "No Edging"/"Raw" → `0`.
- Drill templates (any component named "Drill Template" / "Template") are
  filtered out.
- Width or length showing `0` means the dim wasn't auto-detected — verify
  against the drawing. The component name is shown in the Notes column so
  you can cross-reference.

## How Classification Works (v2)

The tool uses **3D orientation and position** rather than just dimensions:

### Axis Detection
Each panel's thinnest dimension determines its orientation:
- **Thin on X axis** (faces left/right) → Side Panels, End Panels
- **Thin on Y axis** (faces up/down) → Tops, Bases, Shelves, Rails
- **Thin on Z axis** (faces front/back) → Backs, Doors, Plinths

### Position Detection
Within each orientation group, 3D position distinguishes:
- **Tops vs Bases**: Y-position (highest = Top, lowest = Base)
- **Backs vs Doors**: Z-position (rearmost = Back, frontmost = Door)
- **End Panels vs Side Panels**: X-position (extreme edges = End Panel)
- **LH vs RH Side Panels**: relative X-position within cabinet pairs

### Edging Rules
Applied automatically based on classification:
- **LLWW** (all edges): Doors, End Panels, Plinths
- **L** (long edges): Side Panels, Tops, Bases, Shelves, Rails
- **0** (none): Backs

### Additional Checks
- **Face count filter**: Bodies with >30 faces filtered as complex hardware
- **Volume ratio**: Panels where actual volume is <85% of bounding box flagged as non-rectangular
- **Full-width detection**: Panels >2500mm flagged for manual splitting per cabinet

## Sidebar Settings

### Material per Part Type
Assign different materials to each classification (e.g. Oak for doors, White for backs).
This replaces the v1 approach of assigning by thickness only.

### Edgebanding Deductions
Toggle on to reduce panel dimensions by edgeband tape thickness. Useful if the
supplier wants cut size before edging rather than finished size. Deductions apply
per the edging code: L deducts from length edges, LLWW deducts from all four.

## What to Check After Extraction

1. **LH vs RH split** — tool detects this from position but verify against design intent
2. **Full-width panels** — flagged in Notes; split per cabinet if needed
3. **End panel thickness** — may show 18mm from model but need ordering at 25mm
4. **Material assignments** — verify sidebar defaults match the project specification

## Requirements

- Python 3.10+
- SolidWorks STEP files exported as AP214 (File → Save As → STEP AP214)
- Models oriented consistently (cabinets upright, backs against a consistent axis)

## Deploying to Streamlit Cloud (hosted access)

To run the tool as a hosted web app — no local install for end users — deploy
to [Streamlit Community Cloud](https://share.streamlit.io/):

1. Sign in to share.streamlit.io with your GitHub account
2. Click **New app**
3. Pick this repo (`dale-ctrl/wcg-panel-takeoff`), branch `main`, main file `app.py`
4. Click **Deploy**

**Important:** in step 3, click **Advanced settings** before deploying and
set **Python version** to **3.12**. The default (3.14 at time of writing)
doesn't have `cadquery-ocp` wheels — the build will fail. `.python-version`
and `runtime.txt` are in the repo as belt-and-braces but only the dropdown
takes effect on first-deploy.

The first build takes ~5 minutes (cadquery is ~200MB). `packages.txt`
installs the Linux system libraries cadquery needs (libGL, libEGL, etc.).

If `cadquery` ever fails to install, the app still boots — the STEP tab
shows a warning and the **PDF Upload** and **Manual Entry** tabs continue
to work. The cadquery import is lazy so a build failure for it won't break
the whole app.

### Updating the deployed app

Push to `main` and Streamlit Cloud auto-redeploys within a minute.

## Troubleshooting

- **Browser doesn't open**: Navigate to http://localhost:8501
- **cadquery install fails**: Try `pip install cadquery-ocp` first, then `pip install cadquery`
- **STEP won't parse**: Ensure AP214 format. AP203 may strip assembly structure.
- **Classifications wrong**: Check model orientation is consistent. If cabinets are rotated
  in the assembly, axis detection may misclassify — the tool assumes Y=up, Z=front/back.
