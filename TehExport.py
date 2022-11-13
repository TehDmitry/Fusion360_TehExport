from . import config
import os
import adsk.core
#from .lib import fusion360utils as futil

from .FusionFileExport import FusionFileExport

import traceback


app = adsk.core.Application.get()
ui = app.userInterface


# TODO *** Specify the command identity information. ***
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_cmdDialog'
CMD_NAME = 'Teh Export'
CMD_Description = 'A Fusion 360 Add-in Command with a dialog'

# Specify that the command will be promoted to the panel.
IS_PROMOTED = True

# TODO *** Define the location where the command button will be created. ***
# This is done by specifying the workspace, the tab, and the panel, and the
# command it will be inserted beside. Not providing the command to position it
# will insert it at the end.
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidScriptsAddinsPanel'
# COMMAND_BESIDE_ID = 'ScriptsManagerCommand'
COMMAND_BESIDE_ID = ''

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(
    os.path.abspath(__file__)), 'resources', '')


# global set of event handlers to keep them referenced for the duration of the command
handlers = []


def run(context):
    try:
        print('run')
        # stop(context)
        cmd_def = ui.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            print(cmd_def)
            # cmd_def.deleteMe()

        # Create a command Definition.
        cmd_def = ui.commandDefinitions.addButtonDefinition(
            CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

        # Define an event handler for the command created event. It will be called when the button is clicked.
        #futil.add_handler(cmd_def.commandCreated, command_created)

        on_command_created = CommandCreatedHandler()
        cmd_def.commandCreated.add(on_command_created)
        handlers.append(on_command_created)

        # ******** Add a button into the UI so the user can run the command. ********
        # Get the target workspace the button will be created in.
        workspace = ui.workspaces.itemById(WORKSPACE_ID)

        # Get the panel the button will be created in.
        panel = workspace.toolbarPanels.itemById(PANEL_ID)

        # Create the button command control in the UI after the specified existing command.
        control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)

        # Specify if the command is promoted to the main toolbar.
        control.isPromoted = IS_PROMOTED
    except:
        futil.handle_error('run')


def stop(context):
    #ui = None
    try:
        print('stop')
        # Get the various UI elements for this command
        workspace = ui.workspaces.itemById(WORKSPACE_ID)
        panel = workspace.toolbarPanels.itemById(PANEL_ID)
        command_control = panel.controls.itemById(CMD_ID)
        command_definition = ui.commandDefinitions.itemById(CMD_ID)

        # Delete the button command control
        if command_control:
            command_control.deleteMe()

        # Delete the command definition
        if command_definition:
            command_definition.deleteMe()
    except Exception as e:
        if ui:
            ui.messageBox(f'AddIn Stop Failed: {e}')


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
        self.app = app
        self.ui = self.app.userInterface        

    def notify(self, args):
        print('CommandExecuteHandler')

        try:

            command = args.firingEvent.sender
            inputs = command.commandInputs

            values = {}
            for input in inputs:
                if hasattr(input, 'value'):
                    values[input.id] = input.value
                elif hasattr(input, 'listItems'):
                    for listItem in input.listItems:
                        if listItem.isSelected:
                            values[input.id] = listItem.name

            print(values)

            output_path = self.ask_for_output_path()
            if not output_path:
                return

            with FusionFileExport(app, output_path) as total_export:
                total_export.export_step = values['export_step']
                total_export.export_iges = values['export_iges']
                total_export.export_stl = values['export_stl']

                if values['exportType'] == 'Hub':
                    total_export.exportActiveHub()
                elif values['exportType'] == 'Project':
                    total_export.exportCurrentProject()
                elif values['exportType'] == 'File':
                    total_export.exportCurrentFile()
                else:
                    if ui:
                        ui.messageBox('Unknown export type: {}'.format(
                            values['exportType']))

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

    def ask_for_output_path(self, message="Where should we store data?"):
        folder_dialog = self.ui.createFolderDialog()
        folder_dialog.title = message
        dialog_result = folder_dialog.showDialog()
        if dialog_result != adsk.core.DialogResults.DialogOK:
            return None

        output_path = folder_dialog.folder
        return output_path


class CommandDestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            print('CommandDestroyHandler')

            # when the command is done, terminate the script
            # this will release all globals which will remove all event handlers
            # adsk.terminate()
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            command = args.command
            command.isRepeatable = False

            # cmd.helpFile = 'help.html'

            # define the inputs
            inputs = command.commandInputs

            inputs.addImageCommandInput('image', '', 'resources/Icon_128.png')

            # Create radio button group input.
            # radioButtonGroup = inputs.addRadioButtonGroupCommandInput(
            #     'exportType', 'Export type')
            # radioButtonItems = radioButtonGroup.listItems
            # radioButtonItems.add("Hub", False)
            # radioButtonItems.add("Project", False)
            # radioButtonItems.add("File", True)
            # radioButtonGroup.isFullWidth = True

            dropdownInput3 = inputs.addDropDownCommandInput('exportType', 'Export type', adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            dropdown3Items = dropdownInput3.listItems
            dropdown3Items.add('Hub', False, '')
            dropdown3Items.add('Project', False, '')
            dropdown3Items.add('File', True, '')

            inputs.addBoolValueInput(
                'export_step', 'Export step', True, "", True)
            inputs.addBoolValueInput(
                'export_stl', 'Export stl', True, "", False)
            inputs.addBoolValueInput(
                'export_iges', 'Export iges', True, "", False)

            onExecute = CommandExecuteHandler()
            command.execute.add(onExecute)
            handlers.append(onExecute)

            onDestroy = CommandDestroyHandler()
            command.destroy.add(onDestroy)
            handlers.append(onDestroy)
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
