# Batch DXF Export — Fusion 360 Add-In

Select multiple planar faces in Fusion 360 and batch-export them as one or more DXF files.

## Features

- **Per-face export** (default) — each selected face becomes a separate DXF file, named `{BodyName}_{index}.dxf`
- **Single-file export** — all selected faces are projected into one DXF
- Toolbar button added to the UTILITIES tab (Add-Ins panel)
- Filters selection to planar faces only

## Installation

1. Open the Fusion 360 AddIns directory:
   ```
   ./install.sh
   ```
   Or navigate there manually:
   ```
   ~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/
   ```

2. Copy the `BatchDXFExport/` folder into that directory.

3. In Fusion 360, go to **UTILITIES → Add-Ins → Scripts and Add-Ins**, find **BatchDXFExport** in the Add-Ins tab, and click **Run**.

## Usage

1. Click the **Batch DXF Export** button in the UTILITIES tab.
2. Select one or more planar faces in the viewport.
3. Optionally check **"Export all faces into one DXF"** for single-file mode.
4. Click **OK**, then choose an output folder (or filename in single-file mode).
5. A summary dialog confirms what was exported.
