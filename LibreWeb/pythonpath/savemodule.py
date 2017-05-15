# this is a part of LibreWeb project

import pickle


class LibreWebPickle:
    def __init__(self, name="libreweb.data"):
        '''Set your name for saved/created file'''
        self.name = name

    def save(self, my_variable):
        with open(self.name, "wb") as file:
            pickle.dump(my_variable, file)
            file.close()

    def read(self):
        with open(self.name, "rb") as file:
            return pickle.load(file)


def insert_data(docs, doc_name, sheet_name, url, tag, array_nr, cell_address, insert_as):

    if doc_name in docs:
        target = docs[doc_name]
        if sheet_name in target:
            target = target[sheet_name]
            if url in target:
                target = target[url]
                if tag in target:
                    docs[doc_name][sheet_name][url][tag][cell_address] = [array_nr, insert_as]
                else:
                    docs[doc_name][sheet_name][url][tag] = {cell_address: [array_nr, insert_as]}
            else:
                docs[doc_name][sheet_name][url] = {tag: {cell_address: [array_nr, insert_as]}}
        else:
            docs[doc_name][sheet_name] = {url: {tag: {cell_address: [array_nr, insert_as]}}}
    else:
        docs[doc_name] = {sheet_name: {url: {tag: {cell_address: [array_nr, insert_as]}}}}
    return docs
