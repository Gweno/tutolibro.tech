# Created by Super Busy Daddy 16/08/2019
# This program displays 'Hello World!" in cell A1 of the 
# current Calc document.
import msgbox


# get the doc from the scripting context 
# which is made available to all scripts
desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()

# access the active sheet
active_sheet = model.CurrentController.ActiveSheet


def HelloWorld(*args):
    """Write 'Hello World!' in Cell A1"""
    
    # get the doc from the scripting context 
    # which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    
    # access the active sheet
    active_sheet = model.CurrentController.ActiveSheet

    # write 'Hello World' in A1
    active_sheet.getCellRangeByName("A1").String = "Hello World!"
    
def write_my_text(my_text):
    """Write what I want in in Cell A1"""
    
    # ~ # get the doc from the scripting context 
    # ~ # which is made available to all scripts
    # ~ desktop = XSCRIPTCONTEXT.getDesktop()
    # ~ model = desktop.getCurrentComponent()
    
    # ~ # access the active sheet
    # ~ active_sheet = model.CurrentController.ActiveSheet

    # write in A1
    active_sheet.getCellRangeByName("A1").String = my_text

    # ~ message = args[0]
    # ~ myBox = msgbox.MsgBox(XSCRIPTCONTEXT.getComponentContext())
    # ~ myBox.addButton("oK")
    # ~ myBox.renderFromButtonSize()
    # ~ myBox.numberOflines = 2
    # ~ myBox.show(message,0,"Title")

def write_1_to_10():
    
    # ~ # get the doc from the scripting context 
    # ~ # which is made available to all scripts
    # ~ desktop = XSCRIPTCONTEXT.getDesktop()
    # ~ model = desktop.getCurrentComponent()
    
    # ~ # access the active sheet
    # ~ active_sheet = model.CurrentController.ActiveSheet

    for each_number in range(10):
        active_sheet.getCellByPosition(1,each_number).String = each_number + 1

def main(*args):
    """Our main program, that call other functions in the order we want"""

    
    write_my_text("My first macro in Python!")
    write_1_to_10()
