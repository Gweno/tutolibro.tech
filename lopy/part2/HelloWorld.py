# Created by Gweno 16/08/2019 for tutolibro.tech
# This program displays 'Hello World!" in cell A1 of the 
# current Calc document.

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
