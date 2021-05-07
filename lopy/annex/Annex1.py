
# get the doc from the scripting context 
# which is made available to all scripts
desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()

# access the active sheet
active_sheet = model.CurrentController.ActiveSheet

def write_1_to_10(*args):
    """Write what I want in in Cell A1"""
    for each_number in range(10):
        active_sheet.getCellByPosition(1,each_number).Value = each_number + 1
        
