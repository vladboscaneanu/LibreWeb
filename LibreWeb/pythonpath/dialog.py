# This is a part of LibreWeb project


class DialogModel:
    '''A class that create a dialog model'''

    def __init__(self, ctx, dlg_model_props):
        self.model = ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx)
        for key, value in dlg_model_props.items():
            setattr(self.model, key, value)

    def add_elements(self, args):
        '''args = (name=String, type=String, props={}), '''
        for elem in args:
            item = self.model.createInstance('com.sun.star.awt.UnoControl' + elem[1] + "Model")
            for key, value in elem[2].items():
                setattr(item, key, value)
            self.model.insertByName(elem[0], item)


def get_dialog_control(ctx, dialog_model):
    dialog_control = ctx.getServiceManager().createInstance("com.sun.star.awt.UnoControlDialog")
    dialog_control.setModel(dialog_model)
    toolkit = ctx.getServiceManager().createInstance("com.sun.star.awt.ExtToolkit")
    dialog_control.setVisible(False)
    dialog_control.createPeer(toolkit, None)
    return dialog_control
