# parser module of LibreWeb project

from html.parser import HTMLParser


class LibreWebParser(HTMLParser):
    def __init__(self, target_tag):
        ''' Please select a target_tag
        for your web page.
        Don't forget to use the feed() function.'''
        HTMLParser.__init__(self)
        self.targetTag = target_tag
        self.targetFound = False
        self.collectedData = []

    def handle_starttag(self, tag, attrs):
        self.targetFound = False
        if tag == self.targetTag:
            self.targetFound = True

    def handle_endtag(self, tag):
        if tag == self.targetTag:
            self.targetFound = False

    def handle_data(self, data):
        if self.targetFound:
            if isinstance(data, str):
                data = data.strip()
            self.collectedData.append(data)
