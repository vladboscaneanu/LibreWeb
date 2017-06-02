# This is a part of LibreWeb project


def dir_as_string(argument):
    string = ""
    for i in dir(argument):
        string = string + i + " -- "
    return string


def get_save_file(ctx):
    '''Returns the save file URL, if directory not
       exists, it will be created'''
    from settings import save_file_name, save_dir_name
    import unohelper
    import os
    _path = ctx.ServiceManager.createInstance("com.sun.star.util.PathSubstitution")
    user_dir = _path.getSubstituteVariableValue("$(user)")
    user_dir = unohelper.fileUrlToSystemPath(user_dir)
    save_dir = os.path.join(user_dir, save_dir_name)
    if save_dir_name not in os.listdir(user_dir):
        os.mkdir(save_dir)
    return os.path.join(save_dir, save_file_name)


def open_url(url, message_box, default_decoding="utf-8"):
    '''Function that try to open and return a
       valid internet address'''
    import urllib.error
    import urllib.request
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


def start_service(save_file, document, message_box, *args):
    try:
        from savemodule import LibreWebPickle
        from parsermodule import LibreWebParser

        documents = LibreWebPickle(save_file).read()
        doc_title = document.getTitle()

        if doc_title in documents:
            sheets = document.Sheets.ElementNames
            sheets_list = list(documents[doc_title].keys())
            for sheet_name in sheets_list:
                if sheet_name in sheets:
                    sheet = document.Sheets.getByName(sheet_name)
                    url_list = list(documents[doc_title][sheet_name].keys())
                    for url in url_list:
                        result = open_url(url, message_box)
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
