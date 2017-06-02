'''Set the Settings dialog'''
dialog_model_props = {"PositionX": 0, "PositionY": 0,
                      "Width": 400, "Height": 300,
                      "Title": "LibreWeb Settings"}

# dialog items general settings

# DocumentName
DocumentName_props = {"Border": 3, "FontHeight": 15, "PositionX": 10,
                      "PositionY": 0, "Width": 380, "Height": 20,
                      "Align": 1, "VerticalAlign": 1}

# URL TextBox
URLTextBox_props = {"Border": 3, "FontHeight": 15, "PositionX": 10,
                    "PositionY": 30, "Width": 380, "Height": 20}

# Sheet Control
SheetControl_props = {"PositionX": 0, "PositionY": 60, "Width": 200,
                      "Height": 90, "Label": "Sheet Control", "FontHeight": 10}

# Tag Control
TagControl_props = {"PositionX": 200, "PositionY": 60, "Width": 200,
                    "Height": 90, "Label": "Tag Control", "FontHeight": 10}

# SheetsListBox
SheetName_props = {"PositionX": 10, "PositionY": 80, "Width": 180,
                   "Height": 20, "Border": 3,
                   "VerticalAlign": 1, "Align": 1,
                   "FontHeight": 15}

# ElementToSearch
ElementToSearch_props = {"Border": 3, "FontHeight": 15, "PositionX": 210,
                         "PositionY": 80, "Width": 180, "Align": 1,
                         "Height": 20}

# ColumnListBox

Columns_props = {"Border": 3, "FontHeight": 15, "PositionX": 10,
                 "PositionY": 110, "Width": 80, "Height": 20, "Align": 1,
                 "Dropdown": True}

# RowListBox

Rows_props = {"Border": 3, "FontHeight": 15, "PositionX": 110,
              "PositionY": 110, "Width": 80, "Height": 20, "Dropdown": True,
              "Align": 1}

# TagCountListBox

TagCount_props = {"Border": 3, "Align": 1, "Dropdown": True, "FontHeight": 15,
                  "PositionX": 210, "PositionY": 110, "Width": 180,
                  "Height": 20, "Enabled": False}

# TagResult

TagResult_props = {"Label": "Tag Result", "FontHeight": 10, "PositionX": 0,
                   "PositionY": 160, "Width": 400, "Height": 80}

# ResultLabel

ResultLabel_props = {"Label": "No data.", "Align": 1, "FontHeight": 15, "Border": 3,
                     "Width": 380, "Height": 60, "PositionX": 10, "PositionY": 170,
                     "VerticalAlign": 1, "MultiLine": True}

# Control Panel

ControlPanel_props = {"Label": "Control Panel", "FontHeight": 10, "PositionX": 0,
                      "PositionY": 250, "Width": 400, "Height": 50}

# CheckURLButton

CheckURL_props = {"Label": "Check URL", "Align": 1, "FontHeight": 15, "Width": 80,
                  "Height": 20, "PositionX": 10, "PositionY": 270,
                  "VerticalAlign": 1}

# SaveAs

SaveAs_props = {"Label": "Save as:", "FontHeight": 15, "Align": 1, "Border": 0,
                "Width": 50, "Height": 20, "PositionX": 110, "PositionY": 270,
                "VerticalAlign": 1}
# DataType
DataType_props = {"Align": 1, "Dropdown": True, "Width": 120,
                  "Height": 20, "PositionX": 170, "PositionY": 270,
                  "Border": 3, "FontHeight": 15,
                  "StringItemList": ["String", "Value"]}

# SaveButton

SaveButton_props = {"Align": 1, "Label": "Save", "FontHeight": 15, "Width": 80,
                    "Height": 20, "PositionX": 310, "PositionY": 270,
                    "VerticalAlign": 1, "Enabled": True}

# the final variable
dialog_items = (("DocumentName", "FixedText", DocumentName_props),
("URLTextBox", "Edit", URLTextBox_props),
("SheetsControl", "GroupBox", SheetControl_props),
("TagControl", "GroupBox", TagControl_props),
("SheetName", "FixedText", SheetName_props),
("ElementToSearch", "Edit", ElementToSearch_props),
("ColumnListBox", "ListBox", Columns_props),
("RowListBox", "ListBox", Rows_props),
("TagCountListBox", "ListBox", TagCount_props),
("TagResult", "GroupBox", TagResult_props),
("ResultLabel", "FixedText", ResultLabel_props),
("ControlPanel", "GroupBox", ControlPanel_props),
("CheckURLButton", "Button", CheckURL_props),
("SaveAs", "FixedText", SaveAs_props),
("DataTypeListBox", "ListBox", DataType_props),
("SaveButton", "Button", SaveButton_props)
                )
