import uno

ctx = XSCRIPTCONTEXT.getComponentContext()
desktop = XSCRIPTCONTEXT.getDesktop()
document = XSCRIPTCONTEXT.getDocument()
from messagebox import MsgBox

msg_box = MsgBox(desktop)


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
        msg_box, get_save_file(ctx), document,
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
    save_file = get_save_file(ctx)
    start_service(save_file, document, msg_box)


def Get_Support(*args):
    import webbrowser
    from settings import support_forum
    webbrowser.open(support_forum)


# End of script
