from __future__ import with_statement

import adsk.core
import adsk.fusion
import adsk.cam
import traceback

from logging import Logger, FileHandler, Formatter
from threading import Thread

import time
import os
import re
import shutil

max_output_path_length = 230
ignore_already_exported_files = True


class FusionFileExport(object):
    def __init__(self, app, output_path):
        self.app = app
        self.ui = self.app.userInterface
        self.data = self.app.data
        self.documents = self.app.documents
        self.log = Logger("Fusion 360 Total Export")
        self.num_issues = 0
        self.was_cancelled = False
        self.has_cloud_export = False
        self.temp_foler_name = "_temp"
        self.exportignore = ""

        self.export_step = True
        self.export_stl = False
        self.export_iges = False

        self.output_path = output_path
        self.initializeOutputPath()

        self.project_index = 0
        self.projects_count = 0

        self.file_index = 0
        self.files_count = 0

        self.file_exported_count = 0
        self.file_skipped_count = 0

        self.progress_dialog = self.ui.createProgressDialog()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def initializeOutputPath(self):
        if os.path.exists(os.path.join(self.output_path, self.temp_foler_name)):
            dialogResult = self.ui.messageBox("Temp folder {} could contain not actual data. Do you really want to continue?".format(
                self.temp_foler_name), 'Fusion 360 total exporter', adsk.core.MessageBoxButtonTypes.YesNoButtonType, adsk.core.MessageBoxIconTypes.InformationIconType)
            if dialogResult == adsk.core.DialogResults.DialogNo:
                return

        if os.path.exists(os.path.join(self.output_path, 'exportignore.txt')):
            f = open(os.path.join(self.output_path, 'exportignore.txt'))
            self.exportignore = f.read()
            f.close()

        file_handler = FileHandler(os.path.join(
            self.output_path, 'output.log'), 'a', 'utf-8')
        file_handler.setFormatter(
            Formatter(u'%(asctime)s - %(levelname)s - %(message)s'))
        self.log.addHandler(file_handler)

    def showExportResult(self):
        self.progress_dialog.hide()

        if self.was_cancelled:
            self.ui.messageBox("Cancelled!")
        elif self.num_issues > 0:
            self.ui.messageBox("The exporting process ran into {num_issues} issue{english_plurals}. Please check the log for more information".format(
                num_issues=self.num_issues,
                english_plurals="s" if self.num_issues > 1 else ""
            ))
        else:
            self.ui.messageBox("Export finished completely successfully!\n{file_exported_count} files exported, {file_skipped_count} files skipped".format(
                file_skipped_count=self.file_skipped_count,
                file_exported_count=self.file_exported_count
            ))

        if self.has_cloud_export:
            self.ui.messageBox(
                "Wait for cloud export finish before closing Fusion360", "Warning!", 0, 3)  # OK, Warning

        if os.path.exists(os.path.join(self.output_path, self.temp_foler_name)):
            self.ui.messageBox(
                "Please delete the temp foler {} manually".format(self.temp_foler_name))

    def exportActiveHub(self):
        self.log.info("Starting export Active Hub")
        self.ui.messageBox(
            "Searching for and exporting files will take a while, depending on how many files you have.\n\n"
            "You won't be able to do anything else. It has to do everything in the main thread and open and close every file.\n\n"
            "Take an early lunch."
        )

        self.progress_dialog.show("Exporting data!", "", 0, 1, 1)
        self._exportHub(self.data.activeHub)
        self.log.info("Done exporting")
        self.showExportResult()

    def exportCurrentProject(self):
        self.log.info("Starting export Current Project")

        product = self.app.activeProduct
        design = adsk.fusion.Design.cast(product)
        if not design.parentDocument.dataFile:
            self.ui.messageBox("Current file is not saved!", 'Error',
                               adsk.core.MessageBoxButtonTypes.OKButtonType, adsk.core.MessageBoxIconTypes.CriticalIconType)
            return        

        self.project_index = 0
        self.projects_count = 1

        self.progress_dialog.show("Exporting data!", "", 0, 1, 1)
        self._exportProject(design.parentDocument.dataFile.parentProject)
        self.log.info("Done exporting")
        self.showExportResult()

    def exportCurrentFile(self):
        self.log.info("Starting export Current File")

        product = self.app.activeProduct
        design = adsk.fusion.Design.cast(product)
        if not design.parentDocument.dataFile:
            self.ui.messageBox("Current file is not saved!", 'Error',
                               adsk.core.MessageBoxButtonTypes.OKButtonType, adsk.core.MessageBoxIconTypes.CriticalIconType)
            return

        self.file_index = 0
        self.files_count = 1

        self._write_data_file(design.parentDocument.dataFile)

        self.log.info("Done exporting")
        self.showExportResult()

    def _exportAllHubs(self):
        if len(self.data.dataHubs) > 1:
            self.ui.messageBox("The API doesn't support activating a hub and you can only work with the contents of a hub in Fusion when that hub is active. \n\nScript will work only in active hub. \n\nhttps://forums.autodesk.com/t5/fusion-360-api-and-scripts/how-to-select-teams-hub-from-python-api/m-p/10748918")

        all_hubs = self.data.dataHubs
        for hub_index in range(all_hubs.count):
            hub = all_hubs.item(hub_index)
            self._exportHub(hub)
            if self.was_cancelled:
                return

    def _exportHub(self, hub):

        self.log.info("Exporting hub \"{}\"".format(hub.name))
        self.data.activeHub = hub
        time.sleep(1)
        self.log.info("activeHub hub \"{}\"".format(
            self.data.activeHub.name))

        if hub.id == self.data.activeHub.id:
            all_projects = hub.dataProjects
            for project_index in range(all_projects.count):

                project = all_projects.item(project_index)
                self._exportProject(project)

                if self.was_cancelled:
                    return

            self.log.info("Finished exporting hub \"{}\"".format(hub.name))

    def _exportProject(self, project):
        self.log.info(
            "Exporting project \"{}\"".format(project.name))
        self.data.activeProject = project
        time.sleep(1)

        self.log.info("activeProject \"{}\"".format(project.name))

        folder = project.rootFolder

        files = []
        files.extend(self._get_files_for(folder))

        self.progress_dialog.message = "Project: {} of {}\nExporting design %v of %m".format(
            self.project_index + 1,
            self.projects_count
        )
        self.progress_dialog.maximumValue = len(files)
        self.progress_dialog.reset()

        if not files:
            self.log.info("No files to export for this project")
            return

        for file_index in range(len(files)):
            self.app.activeViewport.refresh()
            adsk.doEvents()

            if self.progress_dialog.wasCancelled:
                self.log.info("The process was cancelled!")
                self.was_cancelled = True
                return

            file: adsk.core.DataFile = files[file_index]
            self.progress_dialog.progressValue = file_index + 1
            self._write_data_file(file)

        self.log.info(
            "Finished exporting project \"{}\"".format(project.name))

    def _get_files_for(self, folder):
        files = []
        for file in folder.dataFiles:
            files.append(file)

        for sub_folder in folder.dataFolders:
            files.extend(self._get_files_for(sub_folder))

        return files

    def _write_data_file(self, file: adsk.core.DataFile):
        if file.fileExtension != "f3d" and file.fileExtension != "f3z":
            self.log.info("Not exporting file \"{}\"".format(file.name))
            self.file_skipped_count += 1
            return

        root_folder = self.output_path
        # self.log.info("Exporting file \"{}\"".format(file.name))

        try:
            file_folder = file
            file_folder_path = ""

            while True:
                file_folder = file_folder.parentFolder
                if not file_folder.isRoot:
                    file_folder_path = os.path.join(
                        self._name(file_folder.name), file_folder_path)

                if file_folder.parentFolder is None:
                    break

            parent_project = file_folder.parentProject
            parent_hub = parent_project.parentHub

            file_folder_path = self._take(
                root_folder,
                "Hub {}".format(self._name(parent_hub.name)),
                "{}".format(self._name(parent_project.name)),
                file_folder_path,
                self._name(file.name) + "." + file.fileExtension
            )

            self.log.info("Exporting file \"{}\" to \"{}\"".format(
                file.name, file_folder_path))

            if self.is_ignoring_file(file_folder_path):
                self.log.info(
                    "File \"{}\" found in exportignore.txt".format(file_folder_path))
                return

            if not os.path.exists(file_folder_path):
                self.num_issues += 1
                self.log.exception(
                    "Couldn't make root folder\"{}\"".format(file_folder_path))
                return

        except BaseException as ex:
            self.num_issues += 1
            self.log.exception(
                "Failed while working on \"{}\"".format(file.name), exc_info=ex)
            raise

        file_export_path = os.path.join(file_folder_path, self._name(
            file.name)) + " v" + str(file.versionNumber)
        # fix for names with dots. Fusion trying to interpretate symbols after last dot as extensoin
        file_export_path = file_export_path + ".f3d"

        # only for check. can't pass into Fusion. Not self._name(file.name)
        assembly_export_path = os.path.join(
            file_folder_path, file.name) + ".f3z"
        zip_acrhive_path = file_export_path + "_files"

        if max_output_path_length > 0 and len(file_export_path) > max_output_path_length:
            self.file_skipped_count += 1
            self.log.info(
                "Path is too long. Skip \"{}\"".format(file_export_path))
            return

        is_assembly = file.hasChildReferences  # very slow call ~0.2s
        is_file_export_path_exist = os.path.exists(file_export_path)
        is_assembly_export_path_exist = os.path.exists(assembly_export_path)
        is_zip_acrhive_exist = os.path.exists(zip_acrhive_path + ".zip")

        if ignore_already_exported_files and is_file_export_path_exist and (not is_assembly or is_assembly_export_path_exist) and is_zip_acrhive_exist:
            self.file_skipped_count += 1
            self.log.info(
                "All data files \"{}\" already exists".format(file_export_path))
            return

        document = None
        try:
            document = self.documents.open(file)

            if document is None:
                raise Exception("Documents.open returned None")

            document.activate()

            self.log.info("Writing to \"{}\" \"{}\"".format(
                file_folder_path, file_export_path))

            if not os.path.exists(file_export_path + ".png"):
                self.app.activeViewport.refresh()
                adsk.doEvents()
                self.app.activeViewport.saveAsImageFile(
                    file_export_path + '.png', 512, 512)
                self.check_exported_file(file_export_path + '.png')

            if is_assembly and not is_assembly_export_path_exist:
                self.has_cloud_export = True
                self.log.info("f3z file. executing cloud export into \"{}\"".format(
                    assembly_export_path))
                returnValue = self.app.executeTextCommand(
                    u'data.fileExport f3z "' + file_folder_path + '"')
                self.log.info(
                    "cloud export status: \"{}\"".format(returnValue))
                time.sleep(5)

            if not is_file_export_path_exist:
                fusion_document: adsk.fusion.FusionDocument = adsk.fusion.FusionDocument.cast(
                    document)
                design: adsk.fusion.Design = fusion_document.design
                export_manager: adsk.fusion.ExportManager = design.exportManager

                # Write f3d/f3z file
                options = export_manager.createFusionArchiveExportOptions(
                    file_export_path)
                export_manager.execute(options)
                self.check_exported_file(file_export_path)

                # self._write_component(file_folder_path, design.rootComponent)

            if not os.path.exists(zip_acrhive_path) or os.path.getsize(zip_acrhive_path) < 50:
                fusion_document: adsk.fusion.FusionDocument = adsk.fusion.FusionDocument.cast(
                    document)
                design: adsk.fusion.Design = fusion_document.design
                export_manager: adsk.fusion.ExportManager = design.exportManager

                temp_rootComponent_folder_path = self._take(
                    root_folder,
                    self.temp_foler_name,
                    self._name(file.versionId)
                )

                # if len(os.listdir(temp_rootComponent_folder_path)) > 0:
                #   self.log.info("Using cache files for archive \"{}\" -> \"{}\"".format(file.id, file.name))
                # else:
                self.log.info(
                    "Exporting files for archive \"{}\" -> \"{}\"".format(file.id, file.name))
                self._write_component(
                    temp_rootComponent_folder_path, design.rootComponent)

                shutil.make_archive(zip_acrhive_path, 'zip',
                                    temp_rootComponent_folder_path)
                self.check_exported_file(zip_acrhive_path + '.zip')

            self.log.info("Finished exporting file \"{}\"".format(file.name))

        except BaseException as ex:
            self.num_issues += 1
            self.log.exception(
                "Failed while working on \"{}\"".format(file.name), exc_info=ex)

        finally:
            try:
                if document is not None:
                    document.close(False)
            except BaseException as ex:
                self.num_issues += 1
                self.log.exception(
                    "Failed to close \"{}\"".format(file.name), exc_info=ex)

        self.file_exported_count += 1

    def _write_component(self, component_base_path, component: adsk.fusion.Component):
        # design = component.parentDesign

        output_path = os.path.join(
            component_base_path, self._name(component.name))
        if max_output_path_length > 0 and len(output_path) > max_output_path_length:
            self.num_issues += 1
            self.log.error("Path is too long. Skip \"{}\"".format(output_path))
            return

        output_path = self._take(output_path)

        self.log.info("Writing component \"{}\" to \"{}\"".format(
            component.name, output_path))

        try:
            if self.export_step:
                self._write_step(output_path, component)
            if self.export_stl:
                self._write_stl(output_path, component)
            if self.export_iges:
                self._write_iges(output_path, component)
        except Exception as ex:
            self.num_issues += 1
            self.log.exception("Failed " + output_path, exc_info=ex)

        sketches = component.sketches
        for sketch_index in range(sketches.count):
            sketch = sketches.item(sketch_index)
            self._write_dxf(os.path.join(output_path, sketch.name), sketch)

        occurrences = component.occurrences
        for occurrence_index in range(occurrences.count):
            occurrence = occurrences.item(occurrence_index)
            sub_component = occurrence.component
            sub_path = self._take(component_base_path,
                                  self._name(component.name))
            self._write_component(sub_path, sub_component)

    def _write_step(self, output_path, component: adsk.fusion.Component):
        file_path = output_path + ".stp"

        if self.is_ignoring_file(file_path):
            self.log.info(
                "File \"{}\" found in exportignore.txt".format(file_path))
            return

        if os.path.exists(file_path):
            self.log.info("Step file \"{}\" already exists".format(file_path))
            return

        self.log.info("Writing step file \"{}\"".format(file_path))
        export_manager = component.parentDesign.exportManager

        options = export_manager.createSTEPExportOptions(
            output_path, component)
        export_manager.execute(options)

        self.check_exported_file(file_path)

    def _write_stl(self, output_path, component: adsk.fusion.Component):
        file_path = output_path + ".stl"

        if self.is_ignoring_file(file_path):
            self.log.info(
                "File \"{}\" found in exportignore.txt".format(file_path))
            return

        if os.path.exists(file_path):
            self.log.info("Stl file \"{}\" already exists".format(file_path))
            return

        self.log.info("Writing stl file \"{}\"".format(file_path))
        export_manager = component.parentDesign.exportManager

        try:
            options = export_manager.createSTLExportOptions(
                component, output_path)
            export_manager.execute(options)
        except BaseException as ex:
            self.log.exception(
                "Failed writing stl file \"{}\"".format(file_path), exc_info=ex)

            if component.occurrences.count + component.bRepBodies.count + component.meshBodies.count > 0:
                self.num_issues += 1

        self.check_exported_file(file_path)

        bRepBodies = component.bRepBodies
        meshBodies = component.meshBodies

        if (bRepBodies.count + meshBodies.count) > 0:
            self._take(output_path)
            for index in range(bRepBodies.count):
                body = bRepBodies.item(index)
                self._write_stl_body(os.path.join(
                    output_path, body.name), body)

            for index in range(meshBodies.count):
                body = meshBodies.item(index)
                self._write_stl_body(os.path.join(
                    output_path, body.name), body)

    def _write_stl_body(self, output_path, body):
        file_path = output_path + ".stl"

        if self.is_ignoring_file(file_path):
            self.log.info(
                "File \"{}\" found in exportignore.txt".format(file_path))
            return

        if os.path.exists(file_path):
            self.log.info(
                "Stl body file \"{}\" already exists".format(file_path))
            return

        self.log.info("Writing stl body file \"{}\"".format(file_path))
        export_manager = body.parentComponent.parentDesign.exportManager

        try:
            options = export_manager.createSTLExportOptions(body, file_path)
            export_manager.execute(options)
        except BaseException as ex:
            # Probably an empty model, ignore it
            self.num_issues += 1
            self.log.exception(
                "Probably an empty model \"{}\"".format(file_path), exc_info=ex)
            pass

        self.check_exported_file(file_path)

    def _write_iges(self, output_path, component: adsk.fusion.Component):
        file_path = output_path + ".igs"

        if self.is_ignoring_file(file_path):
            self.log.info(
                "File \"{}\" found in exportignore.txt".format(file_path))
            return

        if os.path.exists(file_path):
            self.log.info("Iges file \"{}\" already exists".format(file_path))
            return

        self.log.info("Writing iges file \"{}\"".format(file_path))

        export_manager = component.parentDesign.exportManager

        options = export_manager.createIGESExportOptions(file_path, component)
        export_manager.execute(options)
        self.check_exported_file(file_path)

    def _write_dxf(self, output_path, sketch: adsk.fusion.Sketch):
        file_path = output_path + ".dxf"

        if self.is_ignoring_file(file_path):
            self.log.info(
                "File \"{}\" found in exportignore.txt".format(file_path))
            return

        if os.path.exists(file_path):
            self.log.info(
                "DXF sketch file \"{}\" already exists".format(file_path))
            return

        self.log.info("Writing dxf sketch file \"{}\"".format(file_path))

        if not sketch.saveAsDXF(file_path):
            self.log.error("Could not saveAsDXF \"{}\"".format(
                sketch.errorOrWarningMessage))

        self.check_exported_file(file_path)

    def _take(self, *path):
        out_path = os.path.join(*path)

        if max_output_path_length > 0 and len(out_path) > max_output_path_length:
            self.num_issues += 1
            self.log.error("Path is too long \"{}\"".format(out_path))

        os.makedirs(out_path, exist_ok=True)
        return out_path

    def _name(self, name):
        name = re.sub('[^a-zA-Z0-9а-яА-ЯЁё_ \.,\[\]+-]', '_', name).strip()

        if name.lower().endswith('.stp') or name.lower().endswith('.stl') or name.lower().endswith('.igs'):
            name = name[0: -4] + "_" + name[-3:]

        if name.lower().endswith('.step'):
            name = name[0: -5] + "_" + name[-4:]

        return name

    def is_ignoring_file(self, file_path):
        for line in self.exportignore.splitlines():
            if line.strip() and line.strip() in file_path:
                return True

        return False

    def check_exported_file(self, file_path):
        if not os.path.exists(file_path):
            self.log.error("Exported file \"{}\" not found".format(file_path))
            self.num_issues += 1
            return False

        return True
