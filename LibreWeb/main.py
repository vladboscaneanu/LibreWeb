#

ctx = XSCRIPTCONTEXT.getComponentContext()
smgr = ctx.ServiceManager
desktop = XSCRIPTCONTEXT.getDesktop()
document = XSCRIPTCONTEXT.getDocument()
from messagebox import MsgBox, INFOBOX, WARNINGBOX, QUERYBOX, BUTTONS_OK_CANCEL, OK

msg_box = MsgBox(desktop)


def Set_Data(*args):
    from dialogsettings import dialog_model_props, dialog_items
    from dialog import DialogModel, get_dialog_control
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
        msg_box, smgr, document,
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

    # add some informations at the runtime moment

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
    from tools import start_service  # get_save_file
    # save_file = get_save_file(ctx, msg_box)
    start_service(smgr, document, msg_box)


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
    from tools import get_local_data
    to_file = get_local_data(ctx)
    if to_file:
        to_file_instance = ImportWebData(ctx, msg_box, to_file)
        to_file_instance.import_web_data()


def Export_Web_Settings(*args):
    from importexport import ExportWebData
    from tools import get_local_data
    move_from = get_local_data(ctx)
    if move_from:
        move_from_instance = ExportWebData(ctx, msg_box, move_from)
        move_from_instance.export_web_data()


def Verify_Update(*args):
    from tools import verify_update
    if not (verify_update(ctx, msg_box)):
        msg_box.show("No update available", "Message", INFOBOX)


def Read_Help(*args):
    ''' Read incorporated help '''
    from tools import get_help_file
    from settings import extension_id, help_file
    get_help_file(smgr, desktop, extension_id, help_file)


def Donate_Paypal(*args):
    from settings import url_paypal
    import webbrowser
    webbrowser.open(url_paypal)


def About(*args):
    '''Getting current version'''
    from tools import get_cur_version
    current_version = get_cur_version(ctx)
    msg_box.show("An internet tool for LibreOffice.\n" +
                 "Current version : " + current_version, "LibreWeb", INFOBOX)


def Send_Email(*args):
    '''Send me an email, using a client'''
    from messagebox import BUTTONS_OK_CANCEL, CANCEL, QUERYBOX
    from dialog import DialogModel, get_dialog_control
    from dialogsendemail import dialog_model_props, dialog_items
    from listeners import SendEmailButtonListener
    if msg_box.show(
            "For a correct work of this feature \n you must have an email client.\n Continue ?",
            "Email client ", QUERYBOX, BUTTONS_OK_CANCEL) == CANCEL:
        return
    # create dialog
    dialog_instance = DialogModel(ctx, dialog_model_props)
    dialog_instance.add_elements(dialog_items)
    dialog_control = get_dialog_control(ctx, dialog_instance.model)
    SendButton_control = dialog_control.getControl("SendButton")
    Subject_control = dialog_control.getControl("Subject")
    Message_control = dialog_control.getControl("Message")
    SendButton_control.addActionListener(
        SendEmailButtonListener(ctx, Subject_control, Message_control, msg_box)
    )
    dialog_control.execute()


def Import_Old_Data(*args):
    import os
    from tools import get_local_data
    old_storage = get_local_data(ctx)
    if os.path.isfile(old_storage):
        from savemodule import LibreWebPickle
        old_instance = LibreWebPickle(old_storage)
        old_data = old_instance.read()
        if document.Title in old_data:
            if msg_box.show("Would you like to move local data to document?",
                            "Please confirm", QUERYBOX, BUTTONS_OK_CANCEL) == OK:
                from docsave import save_file, check_save_file
                if check_save_file(document):
                    if msg_box.show("Document storage data will be rewritten, continue?",
                                    "Please confirm", QUERYBOX, BUTTONS_OK_CANCEL) == OK:
                        save_file(smgr, document, old_data[document.Title])
                        del old_data[document.Title]
                        old_instance.save(old_data)
                        msg_box.show(
                            "Local data was imported correctly", "Message", INFOBOX)

                else:
                    save_file(smgr, document, old_data[document.Title])
                    del old_data[document.Title]
                    old_instance.save(old_data)
                    msg_box.show(
                        "Local data was imported correctly", "Message", INFOBOX)
        else:
            msg_box.show("No local data found for this document",
                         "Message", INFOBOX)
    else:
        msg_box.show("No local data found", "Message", INFOBOX)


def cancel_me(*args):
    from docsave import read_file
    stored_data = read_file(smgr, document)
    msg_box.show(str(stored_data))
# End of script
