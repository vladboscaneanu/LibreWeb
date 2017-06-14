# import uno

ctx = XSCRIPTCONTEXT.getComponentContext()
desktop = XSCRIPTCONTEXT.getDesktop()
document = XSCRIPTCONTEXT.getDocument()
from messagebox import MsgBox,INFOBOX,WARNINGBOX,QUERYBOX
from tools import do_update

msg_box = MsgBox(desktop)

# Check for an eventual update
do_update(ctx, msg_box)


def Settings(*args):
    from dialogsettings import dialog_model_props, dialog_items
    from dialog import DialogModel, get_dialog_control
    from tools import get_save_file
    from listeners import TagCountListener, SaveButtonListener, CheckURLListener

    dlg_instance = DialogModel(ctx, dialog_model_props)
    dlg_instance.add_elements(dialog_items)
    dlg_model = dlg_instance.model
    dlg_control = get_dialog_control(ctx, dlg_model)

    # getting items control

    URLTextBox_control = dlg_control.getControl("URLTextBox")
    ElementToSearch_control = dlg_control.getControl("ElementToSearch")
    TagCount_control = dlg_control.getControl("TagCountListBox")
    CheckURL_control = dlg_control.getControl("CheckURLButton")
    ResultLabel_control = dlg_control.getControl("ResultLabel")
    DocumentName_control = dlg_control.getControl("DocumentName")
    SheetName_control = dlg_control.getControl("SheetName")
    SaveButton_control = dlg_control.getControl("SaveButton")
    ColumnListBox_control = dlg_control.getControl("ColumnListBox")
    RowListBox_control = dlg_control.getControl("RowListBox")
    DataTypeListBox_control = dlg_control.getControl("DataTypeListBox")
    # listeners bloc

    TagCount_control.addActionListener(TagCountListener(
        TagCount_control, ResultLabel_control
    ))
    SaveButton_control.addActionListener(SaveButtonListener(
        msg_box, get_save_file(ctx, msg_box), document,
        ColumnListBox_control, RowListBox_control,
        TagCount_control, ElementToSearch_control,
        URLTextBox_control, DataTypeListBox_control
    ))
    CheckURL_control.addActionListener(CheckURLListener(
        msg_box,
        ElementToSearch_control,
        URLTextBox_control, TagCount_control,
        ResultLabel_control
    ))

    # add some information at the runtime moment

    # -- Current sheet's name
    sheet_name = document.getCurrentController().ActiveSheet.Name
    SheetName_control.setText(sheet_name)
    DocumentName_control.setText(document.getTitle())
    columns_count = document.CurrentController.ActiveSheet.Columns.ElementNames[:40]
    ColumnListBox_control.addItems(columns_count, 0)
    rows_count = [x for x in range(1, 41)]
    RowListBox_control.addItems(rows_count, 0)

    # -- run dialog
    dlg_control.execute()


def Start_Service(*args):
    from tools import get_save_file, start_service
    save_file = get_save_file(ctx, msg_box)
    start_service(save_file, document, msg_box)


def Get_Support(*args):
    import webbrowser
    from settings import support_forum
    webbrowser.open(support_forum)


def Check_SSL(*args):
    '''A function that tries to import ssl module,
        shows error on ImportError'''
    try:
        import ssl
        msg_box.show("It seems SSL module works", "Message", INFOBOX)
    except ImportError:
        msg_box.show("It seems SSL module is broken!", "Attention", WARNINGBOX)


def Import_Web_Settings(*args):
    from importexport import ImportWebData
    from tools import get_save_file
    to_file = get_save_file(ctx, msg_box)
    if to_file:
        to_file_instance = ImportWebData(ctx, msg_box, to_file)
        to_file_instance.import_web_data()


def Export_Web_Settings(*args):
    from importexport import ExportWebData
    from tools import get_save_file
    move_from = get_save_file(ctx, msg_box)
    if move_from:
        move_from_instance = ExportWebData(ctx, msg_box, move_from)
        move_from_instance.export_web_data()


def Verify_Update(*args):
    from tools import _verify_update
    if not (_verify_update(ctx, msg_box)):
        msg_box.show("No update available","Message",INFOBOX)

def Download_Help_File(*args):
    from settings import url_help_file
    import webbrowser
    webbrowser.open(url_help_file)

def About(*args):
    from tools import get_cur_version
    current_version = get_cur_version(ctx)
    msg_box.show("An internet tool for LibreOffice.\n" +
                 "Current version : " + current_version, "LibreWeb", INFOBOX)

# End of script
