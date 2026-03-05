#!/bin/bash
# Opens the Fusion 360 AddIns directory in Finder for easy installation.
# Copy the BatchDXFExport/ folder into this directory, then enable the
# add-in via UTILITIES → Add-Ins → Scripts and Add-Ins in Fusion 360.

ADDINS_DIR="$HOME/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns"

mkdir -p "$ADDINS_DIR"
open "$ADDINS_DIR"
open "."
