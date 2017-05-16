# -- Main script  of LibreWeb project
# -----------------------------------
# -- author Vladimir Boscaneanu
# -- e-mail: vladboscaneanu@gmail.com
# ------------------------------------

# ----importing modules

# Standart modules
import os
import urllib.request
import urllib.error
import webbrowser

# LibreOffice modules
import uno
import unohelper
from com.sun.star.awt import XActionListener

# LibreWeb modules

from  savemodule import LibreWebPickle
from savemodule import insert_data
from messagebox import MsgBox
from parsermodule import LibreWebParser
import settings

# ----Global variables

ctx = XSCRIPTCONTEXT.getComponentContext()
desktop = XSCRIPTCONTEXT.getDesktop()
document = XSCRIPTCONTEXT.getDocument()
dialog_title = "LibreWeb Settings"
doc_title = document.getTitle()
dialog_width = 400
dialog_height = 300
message_box = MsgBox(desktop)
dialog_x = 0
dialog_y = 0
accepted_tags = settings.tags
collected_items = []
save_file_name = "libreweb.data"
save_dir_name = "LibreWebSettings"

# getting the user folder

_path = ctx.ServiceManager.createInstance("com.sun.star.util.PathSubstitution")
user_dir = _path.getSubstituteVariableValue("$(user)")
user_dir = unohelper.fileUrlToSystemPath(user_dir)

# looking for LibreWeb folder, if not exists let's create
save_dir = os.path.join(user_dir, save_dir_name)

if save_dir_name not in os.listdir(user_dir):
    os.mkdir(save_dir)
else:
    save_file_url = os.path.join(save_dir, save_file_name)

# Create a pickle instance to work with save_file

save_file = LibreWebPickle(save_file_url)


# ------ Defining functions

def get_support(*args):
    webbrowser.open(settings.support_forum)


def create_istance(name):
    return ctx.getServiceManager().createInstanceWithContext(name, ctx)


def add_to_dialog(name, type, props={}):
    model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
    dialog_model.insertByName(name, model)
    for key, value in props.items():
        setattr(model, key, value)


def open_url(url, default_decoding="utf-8"):
    try:
        result = urllib.request.urlopen(url).read().decode(default_decoding)
        return result
    except UnicodeDecodeError:
        message_box.show("Site is not utf-8 encoded", "Error", 2)

    except ValueError as error:
        message_box.show("Invalid URL", "Attention", 2)

    except urllib.error.URLError:
        message_box.show("Site not reachable or internet connection is missing",
                         "Attention", 2)

    except Exception as error:
        str_error = ""
        for i in error.args:
            str_error = str_error + "\n" + str(i)
            message_box.show(str_error, "Attention", 2)


# -- Listeners

# -- Check Url Button


class CheckURLListener(unohelper.Base, XActionListener):
    def actionPerformed(self, event):
        target_tag = search_element_control.Text.strip()
        target_url = URLTextBox_control.Text.strip()
        global collected_items
        # -- clean previous items
        if collected_items:
            TagCount_control.removeItems(0, len(collected_items))
            TagCount_control.setEnable(False)
            ResultLabel_control.setText("No data.")
        if target_tag in accepted_tags:
            result = open_url(target_url)
            if result:
                my_parser = LibreWebParser(target_tag)
                my_parser.feed(result)
                collected_items = my_parser.collectedData
                if collected_items:
                    TagCount_control.addItems([x for x in range(0, len(collected_items))], 0)
                    TagCount_control.setEnable(True)
        else:
            message_box.show("Not supported tag!", "Attention!", 2)

    def disposing(self, event):
        pass


# -- TagCountListBox


class TagCountListener(unohelper.Base, XActionListener):
    def actionPerformed(self, event):
        item_position = TagCount_control.getSelectedItemPos()
        if item_position >= 0:
            ResultLabel_control.setText(collected_items[item_position])

    def disposing(self, event):
        pass


# -- SaveButtonListener
class SaveButtonListener(unohelper.Base, XActionListener):
    def actionPerformed(self, event):
        # Check if file for save data exists,otherwise let's create

        if not os.path.isfile(save_file_url):
            try:
                save_file.save(settings.documents)
            except OSError as error:
                if error.errno == 13:
                    message_box.show("You have no rights to create the save file",
                                     "Attention", 2)
        if os.path.isfile(save_file_url):
            try:
                if ColumnListBox_control.getSelectedItemPos() == -1:
                    message_box.show("Please select a column", "Error", 2)
                elif RowListBox_control.getSelectedItemPos() == -1:
                    message_box.show("Please select  a row", "Error", 2)
                elif TagCount_control.getSelectedItemPos() == -1:
                    message_box.show("Please select a tag number", "Error", 2)
                elif search_element_control.Text.strip() not in accepted_tags:
                    message_box.show("Not supported tag", "Attention", 2)
                elif not open_url(URLTextBox_control.Text.strip()):
                    pass
                elif DataTypeListBox_control.getSelectedItemPos() == -1:
                    message_box.show("Please select type of inserted data",
                                     "Attention", 2)
                else:
                    # start to save our data
                    docs = save_file.read()
                    sheet_name = document.getCurrentController().ActiveSheet.Name
                    url = URLTextBox_control.Text
                    tag = search_element_control.Text
                    array_nr = TagCount_control.getSelectedItemPos()
                    column = str(ColumnListBox_control.getSelectedItem())
                    row = str(RowListBox_control.getSelectedItem())
                    cell_address = column + row
                    insert_as = DataTypeListBox_control.getSelectedItem()
                    # insert_data is a function, that create our new row in savefile
                    new_docs = insert_data(docs, doc_title, sheet_name, url, tag, array_nr, cell_address, insert_as)
                    # then save all data
                    save_file.save(new_docs)
                    message_box.show("Operation Completed!", "Well done", 1)
            except OSError as error:
                if error.errno == 13:
                    message_box.show("You have no rights to insert new data",
                                     "Attention", 2)

    def disposing(self, event):
        pass


# -- create dialog model
dialog_model = create_istance("com.sun.star.awt.UnoControlDialogModel")
dialog_model_props = {"PositionX": dialog_x, "PositionY": dialog_y,
                      "Width": dialog_width, "Height": dialog_height,
                      "Title": dialog_title}
for key, value in dialog_model_props.items():
    setattr(dialog_model, key, value)

# -- adding dialog elements
# -- Document Name

add_to_dialog("DocumentName", "FixedText", settings.DocumentName_props)

# URL TextBox

add_to_dialog("URLTextBox", "Edit", settings.URLTextBox_props)

# Sheet Control

add_to_dialog("Sheet Control", "GroupBox", settings.SheetControl_props)

# Tag Control

add_to_dialog("TagControl", "GroupBox", settings.TagControl_props)

# SheetsListBox

add_to_dialog("SheetName", "FixedText", settings.SheetName_props)

# ElementToSearch

add_to_dialog("ElementToSearch", "Edit", settings.ElementToSearch_props)

# ColumnListBox

add_to_dialog("ColumnListBox", "ListBox", settings.Columns_props)

# RowListBox

add_to_dialog("RowListBox", "ListBox", settings.Rows_props)

# TagCountListBox

add_to_dialog("TagCountListBox", "ListBox", settings.TagCount_props)

# TagResult

add_to_dialog("TagResult", "GroupBox", settings.TagResult_props)

# ResultLabel

add_to_dialog("ResultLabel", "FixedText", settings.ResultLabel_props)

# Control Panel

add_to_dialog("ControlPanel", "GroupBox", settings.ControlPanel_props)

# CheckURLButton

add_to_dialog("CheckURLButton", "Button", settings.CheckURL_props)

# SaveAs

add_to_dialog("SaveAs", "FixedText", settings.SaveAs_props)

# DataType

add_to_dialog("DataTypeListBox", "ListBox", settings.DataType_props)

# SaveButton

add_to_dialog("SaveButton", "Button", settings.SaveButton_props)

# -- create a dialog control

dialog = create_istance("com.sun.star.awt.UnoControlDialog")
dialog.setModel(dialog_model)

# -- get the items control
URLTextBox_control = dialog.getControl("URLTextBox")
search_element_control = dialog.getControl("ElementToSearch")
TagCount_control = dialog.getControl("TagCountListBox")
CheckURL_control = dialog.getControl("CheckURLButton")
ResultLabel_control = dialog.getControl("ResultLabel")
DocumentName_control = dialog.getControl("DocumentName")
SheetName_control = dialog.getControl("SheetName")
SaveButton_control = dialog.getControl("SaveButton")
ColumnListBox_control = dialog.getControl("ColumnListBox")
RowListBox_control = dialog.getControl("RowListBox")
DataTypeListBox_control = dialog.getControl("DataTypeListBox")

# -- add listeners
CheckURL_control.addActionListener(CheckURLListener())
TagCount_control.addActionListener(TagCountListener())

SaveButton_control.addActionListener(SaveButtonListener())
# -- show dialog

toolkit = create_istance("com.sun.star.awt.ExtToolkit")
dialog.setVisible(False)
dialog.createPeer(toolkit, None)


# -- Main function


def Settings(*args):
    # add some information at the runtime moment

    # -- Current sheet's name
    Sheet_name = document.getCurrentController().ActiveSheet.Name
    SheetName_control.setText(Sheet_name)
    DocumentName_control.setText(doc_title)
    Columns_count = document.CurrentController.ActiveSheet.Columns.ElementNames[:40]
    dialog_model.getByName("ColumnListBox").StringItemList = Columns_count
    Rows_count = [x for x in range(1, 41)]
    dialog_model.getByName("RowListBox").StringItemList = Rows_count

    # -- run dialog
    dialog.execute()


def Start_service(*args):
    try:
        documents = save_file.read()

        if doc_title in documents:
            sheets = document.Sheets.ElementNames
            sheets_list = list(documents[doc_title].keys())
            for sheet_name in sheets_list:
                if sheet_name in sheets:
                    sheet = document.Sheets.getByName(sheet_name)
                    url_list = list(documents[doc_title][sheet_name].keys())
                    for url in url_list:
                        result = open_url(url)
                        if result:
                            tag_list = list(documents[doc_title][sheet_name][url].keys())
                            for tag in tag_list:
                                my_parser = LibreWebParser(tag)
                                my_parser.feed(result)
                                if my_parser.collectedData:
                                    cell_list = list(documents[doc_title][sheet_name][url][tag].keys())
                                    for cell in cell_list:
                                        cell_items = documents[doc_title][sheet_name][url][tag][cell]
                                        cell_data = my_parser.collectedData[cell_items[0]]
                                        if cell_items[1] == "String":
                                            sheet.getCellRangeByName(cell).setString(cell_data)
                                        elif cell_items[1] == "Value":
                                            sheet.getCellRangeByName(cell).setValue(cell_data)
                else:
                    message_box.show("No valid sheets names found", "Error", 2)
                    return

            message_box.show("Operation completed.", "Message", 1)

        else:
            message_box.show("This document has no saved data,go to Settings", "Attention", 1)

    except FileNotFoundError:
        message_box.show("No settings file detected,first save some data!", "Attention", 2)


# End of script
g_ExportedScripts = Start_service
