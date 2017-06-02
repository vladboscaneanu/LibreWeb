import os

import unohelper
from com.sun.star.awt import XActionListener
from settings import _collected_items

collected_items = _collected_items


class TagCountListener(unohelper.Base, XActionListener):
    def __init__(self, tag_count_control, result_label_control):
        self.TagCount_control = tag_count_control
        self.ResultLabel_control = result_label_control

    def actionPerformed(self, event):
        item_position = self.TagCount_control.getSelectedItemPos()
        if item_position >= 0:
            global collected_items
            self.ResultLabel_control.setText(collected_items[item_position])

    def disposing(self, event):
        pass


# -- SaveButtonListener
class SaveButtonListener(unohelper.Base, XActionListener):
    def __init__(self, msg_box, save_file, document,
                 ColumnListBox_ctrl, RowListBox_ctrl,
                 TagCount_ctrl, ElementToSearch_ctrl,
                 URLTextBox_ctrl, DataTypeListBox_ctrl):
        from tools import open_url
        from savemodule import LibreWebPickle, insert_data
        from settings import tags

        self.open_url = open_url
        self.insert_data = insert_data
        self.doc = document
        self.ColumnListBox_ctrl = ColumnListBox_ctrl
        self.RowList_ctrl = RowListBox_ctrl
        self.TagCount_ctrl = TagCount_ctrl
        self.ElementToSearch_ctrl = ElementToSearch_ctrl
        self.URLTextBox_ctrl = URLTextBox_ctrl
        self.DataTypeListBox_ctrl = DataTypeListBox_ctrl
        self.sv_file = save_file

        self.save_ins = LibreWebPickle(self.sv_file)
        self.msg_box = msg_box

        self.tags = tags

    def actionPerformed(self, event):
        # Check if file for save data exists,otherwise let's create

        if not os.path.isfile(self.sv_file):
            try:
                self.save_ins.save({})
            except OSError as error:
                if error.errno == 13:
                    self.msg_box.show("You have no rights to create the save file",
                                      "Attention", 2)
        if os.path.isfile(self.sv_file):
            try:
                if self.ColumnListBox_ctrl.getSelectedItemPos() == -1:
                    self.msg_box.show("Please select a column", "Error", 2)
                elif self.RowList_ctrl.getSelectedItemPos() == -1:
                    self.msg_box.show("Please select  a row", "Error", 2)
                elif self.TagCount_ctrl.getSelectedItemPos() == -1:
                    self.msg_box.show("Please select a tag number", "Error", 2)
                elif self.ElementToSearch_ctrl.Text.strip() not in self.tags:
                    self.msg_box.show("Not supported tag", "Attention", 2)
                elif not self.open_url(self.URLTextBox_ctrl.Text.strip(), self.msg_box):
                    pass
                elif self.DataTypeListBox_ctrl.getSelectedItemPos() == -1:
                    self.msg_box.show("Please select type of inserted data",
                                      "Attention", 2)
                else:
                    # start to save our data
                    docs = self.save_ins.read()
                    sheet_name = self.doc.getCurrentController().ActiveSheet.Name
                    url = self.URLTextBox_ctrl.Text
                    tag = self.ElementToSearch_ctrl.Text
                    array_nr = self.TagCount_ctrl.getSelectedItemPos()
                    column = str(self.ColumnListBox_ctrl.getSelectedItem())
                    row = str(self.RowList_ctrl.getSelectedItem())
                    cell_address = column + row
                    insert_as = self.DataTypeListBox_ctrl.getSelectedItem()
                    # insert_data is a function, that create our new entry in savefile
                    doc_title = self.doc.getTitle()
                    new_docs = self.insert_data(docs, doc_title, sheet_name, url, tag, array_nr, cell_address,
                                                insert_as)
                    # then save all data
                    self.save_ins.save(new_docs)
                    self.msg_box.show("Operation Completed!", "Well done", 1)
            except OSError as error:
                if error.errno == 13:
                    self.msg_box.show("You have no rights to insert new data",
                                      "Attention", 2)

    def disposing(self, event):
        pass


class CheckURLListener(unohelper.Base, XActionListener):
    def __init__(self, msg_box,
                 ElementToSearch_ctrl, URLTextBox_ctrl,
                 TagCount_ctrl, ResultLabel_ctrl):
        from settings import tags
        self.tags = tags
        self.msg_box = msg_box
        self.URLTextBox_ctrl = URLTextBox_ctrl
        self.ElementToSearch_ctrl = ElementToSearch_ctrl
        self.TagCount_ctrl = TagCount_ctrl
        self.ResultLabel_ctrl = ResultLabel_ctrl

    def actionPerformed(self, event):
        from tools import open_url
        target_tag = self.ElementToSearch_ctrl.Text.strip()
        target_url = self.URLTextBox_ctrl.Text.strip()
        global collected_items
        # -- clean previous items
        if collected_items:
            self.TagCount_ctrl.removeItems(0, len(collected_items))
            self.TagCount_ctrl.setEnable(False)
            self.ResultLabel_ctrl.setText("No data.")
        if target_tag in self.tags:
            result = open_url(target_url, self.msg_box)
            if result:
                from parsermodule import LibreWebParser
                my_parser = LibreWebParser(target_tag)
                my_parser.feed(result)
                collected_items = my_parser.collectedData
                if collected_items:
                    self.TagCount_ctrl.addItems([x for x in range(0, len(collected_items))], 0)
                    self.TagCount_ctrl.setEnable(True)
        else:
            self.msg_box.show("Not supported tag!", "Attention!", 2)

    def disposing(self, event):
        pass
