import adsk.core
import adsk.fusion
import os
import re
import traceback

_app = None
_ui = None

# Persistent references to event handlers so they don't get garbage-collected.
_handlers = []

CMD_ID = "batchDxfExportCmd"
CMD_NAME = "Batch DXF Export"
CMD_DESCRIPTION = "Select planar faces and batch-export them as DXF files."
TOOLBAR_TAB_ID = "ToolsTab"
TOOLBAR_PANEL_ID = "SolidScriptsAddinsPanel"


def run(context):
    global _app, _ui
    try:
        _app = adsk.core.Application.get()
        _ui = _app.userInterface

        cmd_defs = _ui.commandDefinitions
        # Clean up any leftover definition from a previous session.
        existing = cmd_defs.itemById(CMD_ID)
        if existing:
            existing.deleteMe()

        cmd_def = cmd_defs.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESCRIPTION)

        on_created = _CommandCreatedHandler()
        cmd_def.commandCreated.add(on_created)
        _handlers.append(on_created)

        # Add the button to the UTILITIES tab → Add-Ins panel.
        workspace = _ui.workspaces.itemById("FusionSolidEnvironment")
        panel = None
        if workspace:
            tabs = workspace.toolbarTabs
            utilities_tab = tabs.itemById(TOOLBAR_TAB_ID)
            if utilities_tab:
                panel = utilities_tab.toolbarPanels.itemById(TOOLBAR_PANEL_ID)
                if not panel:
                    panel = utilities_tab.toolbarPanels.add(TOOLBAR_PANEL_ID, "Batch DXF Export")

        # Fallback to global panel lookup.
        if not panel:
            panel = _ui.allToolbarPanels.itemById(TOOLBAR_PANEL_ID)

        if panel:
            existing_ctrl = panel.controls.itemById(CMD_ID)
            if not existing_ctrl:
                panel.controls.addCommand(cmd_def)

    except Exception:
        if _ui:
            _ui.messageBox(f"Failed to start BatchDXFExport:\n{traceback.format_exc()}")


def stop(context):
    try:
        # Mirror the same lookup used in run().
        panel = None
        workspace = _ui.workspaces.itemById("FusionSolidEnvironment")
        if workspace:
            tabs = workspace.toolbarTabs
            utilities_tab = tabs.itemById(TOOLBAR_TAB_ID)
            if utilities_tab:
                panel = utilities_tab.toolbarPanels.itemById(TOOLBAR_PANEL_ID)
        if not panel:
            panel = _ui.allToolbarPanels.itemById(TOOLBAR_PANEL_ID)

        if panel:
            ctrl = panel.controls.itemById(CMD_ID)
            if ctrl:
                ctrl.deleteMe()

        cmd_def = _ui.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            cmd_def.deleteMe()

        _handlers.clear()
    except Exception:
        if _ui:
            _ui.messageBox(f"Failed to stop BatchDXFExport:\n{traceback.format_exc()}")


# ---------------------------------------------------------------------------
# Command Created — build the dialog inputs
# ---------------------------------------------------------------------------

class _CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        try:
            cmd = adsk.core.Command.cast(args.command)
            inputs = cmd.commandInputs

            # Face selection input — planar faces only, at least one required.
            sel_input = inputs.addSelectionInput(
                "faceSelection", "Faces", "Select planar faces to export"
            )
            sel_input.addSelectionFilter("PlanarFaces")
            sel_input.setSelectionLimits(1, 0)  # min 1, no max

            # Single-file checkbox.
            inputs.addBoolValueInput(
                "singleFile", "Export all faces into one DXF", True, "", False
            )

            on_execute = _CommandExecuteHandler()
            cmd.execute.add(on_execute)
            _handlers.append(on_execute)

        except Exception:
            if _ui:
                _ui.messageBox(f"CommandCreated failed:\n{traceback.format_exc()}")


# ---------------------------------------------------------------------------
# Command Execute — export the DXFs
# ---------------------------------------------------------------------------

class _CommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        try:
            cmd = adsk.core.Command.cast(args.command)
            inputs = cmd.commandInputs

            sel_input = inputs.itemById("faceSelection")
            single_file = inputs.itemById("singleFile").value

            # Collect selected faces.
            faces = []
            for i in range(sel_input.selectionCount):
                entity = sel_input.selection(i).entity
                face = adsk.fusion.BRepFace.cast(entity)
                if face:
                    faces.append(face)

            if not faces:
                _ui.messageBox("No faces selected.")
                return

            if single_file:
                _export_single_file(faces)
            else:
                _export_per_face(faces)

        except Exception:
            if _ui:
                _ui.messageBox(f"Export failed:\n{traceback.format_exc()}")


# ---------------------------------------------------------------------------
# Export modes
# ---------------------------------------------------------------------------

def _export_per_face(faces):
    """Export each face as a separate DXF. Prompts for an output folder."""
    dlg = _ui.createFolderDialog()
    dlg.title = "Choose output folder for DXF files"
    result = dlg.showDialog()
    if result != adsk.core.DialogResults.DialogOK:
        return

    folder = dlg.folder

    exported = 0
    errors = []
    for idx, face in enumerate(faces):
        body = face.body
        body_name = sanitize_filename(body.name) if body else "Face"
        filename = f"{body_name}_{idx}.dxf"
        filepath = os.path.join(folder, filename)

        ok, err = export_face_as_dxf(face, filepath)
        if ok:
            exported += 1
        else:
            errors.append(f"{filename}: {err}")

    msg = f"Exported {exported} of {len(faces)} face(s) to:\n{folder}"
    if errors:
        msg += "\n\nErrors:\n" + "\n".join(errors)
    _ui.messageBox(msg)


def _export_single_file(faces):
    """Export all faces into a single DXF. Prompts for a save-file path."""
    dlg = _ui.createFileDialog()
    dlg.title = "Save combined DXF"
    dlg.filter = "DXF files (*.dxf)"
    dlg.isMultiSelectEnabled = False
    result = dlg.showSave()
    if result != adsk.core.DialogResults.DialogOK:
        return

    filepath = dlg.filename
    if not filepath.lower().endswith(".dxf"):
        filepath += ".dxf"

    design = adsk.fusion.Design.cast(_app.activeProduct)
    root_comp = design.rootComponent
    xy_plane = root_comp.xYConstructionPlane

    sketch = root_comp.sketches.add(xy_plane)
    sketch.name = "_BatchDXF_temp"

    errors = []
    for face in faces:
        try:
            sketch.project(face)
        except Exception:
            try:
                sketch.project2(face, False)
            except Exception as e:
                errors.append(str(e))

    try:
        sketch.saveAsDXF(filepath)
    except Exception:
        try:
            export_mgr = design.exportManager
            opts = export_mgr.createDXFSketchExportOptions(filepath, sketch)
            export_mgr.execute(opts)
        except Exception as e:
            errors.append(f"saveAsDXF failed: {e}")

    sketch.deleteMe()

    msg = f"Exported combined DXF to:\n{filepath}"
    if errors:
        msg += "\n\nWarnings:\n" + "\n".join(errors)
    _ui.messageBox(msg)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def export_face_as_dxf(face, filepath):
    """Create a temp sketch on *face*, project its edges, save as DXF, clean up.

    Returns (success: bool, error_message: str | None).
    """
    design = adsk.fusion.Design.cast(_app.activeProduct)
    comp = face.body.parentComponent

    sketch = comp.sketches.add(comp.xYConstructionPlane)
    try:
        try:
            sketch.project(face)
        except Exception:
            sketch.project2(face, False)

        try:
            sketch.saveAsDXF(filepath)
        except Exception:
            export_mgr = design.exportManager
            opts = export_mgr.createDXFSketchExportOptions(filepath, sketch)
            export_mgr.execute(opts)

        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        sketch.deleteMe()


def sanitize_filename(name):
    """Strip characters that are unsafe in file names."""
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()
