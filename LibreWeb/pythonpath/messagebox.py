# This is a part of LibreWeb project
###################################
#import uno

from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, \
                                               BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, \
                                               BUTTONS_ABORT_IGNORE_RETRY


# initialisate class


class MsgBox():
    def __init__(self, desktop):
        '''As an argument, give the XSCRIPTCONTEXT.getDesktop()'''

        self.parentwin = desktop.getCurrentComponent()
        self.parentwin = self.parentwin.CurrentController
        self.parentwin = self.parentwin.Frame.ContainerWindow
        self.toolkit = self.parentwin.getToolkit()

    def show(self, boxText="Your text", boxTitle="Your title", boxType=MESSAGEBOX, buttonType=BUTTONS_OK):
        '''Change to your pleasure boxType, buttonType, boxTitle and boxText '''

        box = self.toolkit.createMessageBox(self.parentwin, boxType, buttonType, boxTitle, boxText)
        return box.execute()
