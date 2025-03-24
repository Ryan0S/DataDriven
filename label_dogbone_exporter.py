"""This file acts as the main module for this script."""

import traceback
import adsk.core
import adsk.fusion
# import adsk.cam

# Initialize the global variables for the Application and UserInterface objects.
app = adsk.core.Application.get()
ui  = app.userInterface


def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        # Paths
        template_path = r'C:\Users\Ryan\Downloads\Dogbone.f3d'
        export_folder = r'C:\Users\Ryan\Downloads\DogboneExports'  # Customize export folder

        # List of tag IDs to generate
        tag_ids = ['001', '020', '300']

        for tag in tag_ids:
            # Open the dogbone template (local .f3d file)
            import_mgr = app.importManager
            f3d_options = import_mgr.createFusionArchiveImportOptions(template_path)
            doc = import_mgr.importToNewDocument(f3d_options)
            design = adsk.fusion.Design.cast(app.activeProduct)


            # Look for the sketch text labeled "LBL"
            sketch_text_updated = False
            for sketch in design.rootComponent.sketches:
                for sketch_text in sketch.sketchTexts:
                    if sketch_text.text == 'LBL':
                        sketch_text.text = tag
                        sketch_text_updated = True
                        break  # Stop after finding the first match
                if sketch_text_updated:
                    break

            if not sketch_text_updated:
                ui.messageBox(f'Could not find sketch text "LBL" for tag {tag}')
                doc.close(False)
                continue

            # Export the file as .3mf
            export_mgr = design.exportManager
            export_path = f'{export_folder}\\{tag}.3mf'
            options = export_mgr.createC3MFExportOptions(design.rootComponent, export_path)
            export_mgr.execute(options)
            ui.messageBox(f'Successfully exported: {export_path}')



            # Close the document without saving changes
            doc.close(False)

    except Exception as e:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
