"""This file acts as the main module for this script."""

import traceback
import adsk.core
import adsk.fusion

# === CONFIGURATION ===
TEMPLATE_PATH = r'C:\Users\Ryan\Downloads\D638_TYPE_II_LBL.f3d'
EXPORT_FOLDER = r'C:\Users\Ryan\Downloads\DogboneExports'  # Customize export folder
START_ID = 1   # Inclusive
END_ID = 15    # Inclusive (will generate 001 to 007)

# === MAIN FUNCTION ===
def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        # Generate padded IDs
        tag_ids = [str(i).zfill(3) for i in range(START_ID, END_ID + 1)]

        for tag in tag_ids:
            # Open template
            import_mgr = app.importManager
            f3d_options = import_mgr.createFusionArchiveImportOptions(TEMPLATE_PATH)
            doc = import_mgr.importToNewDocument(f3d_options)
            design = adsk.fusion.Design.cast(app.activeProduct)

            # Replace sketch text "LBL"
            sketch_text_updated = False
            for sketch in design.rootComponent.sketches:
                for sketch_text in sketch.sketchTexts:
                    if sketch_text.text == 'LBL':
                        sketch_text.text = tag
                        sketch_text_updated = True
                        break
                if sketch_text_updated:
                    break

            if not sketch_text_updated:
                ui.messageBox(f'❌ Could not find sketch text "LBL" for tag {tag}')
                doc.close(False)
                continue

            # Export .3mf
            export_mgr = design.exportManager
            export_path = f'{EXPORT_FOLDER}\\{tag}.3mf'
            options = export_mgr.createC3MFExportOptions(design.rootComponent, export_path)
            export_mgr.execute(options)
            ui.messageBox(f'✅ Successfully exported: {export_path}')

            # Close without saving
            doc.close(False)

    except Exception as e:
        if ui:
            ui.messageBox(f'Failed:\n{traceback.format_exc()}')
