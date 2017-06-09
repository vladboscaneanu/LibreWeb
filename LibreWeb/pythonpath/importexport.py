class ImportWebData:
    '''Import web settings'''

    def __init__(self, ctx, msgbox, move_to):
        from settings import save_file_name
        self.msg_box = msgbox
        self.destination = move_to
        self.dialog = ctx.ServiceManager.createInstance("com.sun.star.ui.dialogs.FilePicker")
        self.dialog.appendFilter("A LibreWeb file", save_file_name)
        self.dialog.setMultiSelectionMode(False)

    def import_web_data(self):
        '''Shows the dialog ,pick a file, then
           import entire file.'''
        self.dialog.execute()
        file = self.dialog.getSelectedFiles()
        if file:
            from shutil import copyfile
            import os.path
            from unohelper import fileUrlToSystemPath
            file_url = fileUrlToSystemPath(file[0])
            try:
                if os.path.isfile(self.destination):
                    answer = self.msg_box.show("The file exists, replace?", "Warning", 2)
                    if answer:
                        copyfile(file_url, self.destination)
                        self.msg_box.show("Web settings imported successfully", "Message", 1)
            except OSError as error:
                if error.errno == 13:
                    self.msg_box.show("You have no rights to create settings file",
                                      "Attention, 2")


class ExportWebData:
    '''Export web settings'''

    def __init__(self, ctx, msgbox, move_from):
        self.move_from = move_from
        self.msg_box = msgbox
        self.dialog = ctx.ServiceManager.createInstance("com.sun.star.ui.dialogs.FolderPicker")

    def export_web_data(self):
        '''Pick a directory,then export web settings'''
        self.dialog.execute()
        folder = self.dialog.getDirectory()
        if folder:
            from shutil import copy
            import os.path
            from unohelper import fileUrlToSystemPath
            from settings import save_file_name
            folder_url = fileUrlToSystemPath(folder)
            file = os.path.join(folder_url, save_file_name)
            try:
                if os.path.isfile(file):
                    answer = self.msg_box.show("The file exists, replace?", "Warning", 2)
                    if answer:
                        copy(self.move_from, folder_url)
                        self.msg_box.show("Web settings exported successfully", "Message", 1)
                else:
                    copy(self.move_from, folder_url)
                    self.msg_box.show("Web settings exported successfully", "Message", 1)
            except OSError as error:
                if error.errno == 13:
                    self.msg_box.show("You have no rights to save here file ",
                                      "Attention", 2)
