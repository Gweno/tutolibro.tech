# Created by Gwenole Capp 05/05/2021
# For tutolibro.tech
# Public Domain, feel free to copy, modify, use in your own scripts
# 
# email: gwenole.capp@gmail.com
   
# set global variables for context
    
# get the doc from the scripting context 
# which is made available to all scripts
desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()

# access the active sheet
active_sheet = model.CurrentController.ActiveSheet

# define useful functions

def getSelectionAddresses(horizontalOffset = 0 , verticalOffset = 0):
    # get the range of addresses from selection
    oSelection = model.getCurrentSelection()
    oArea = oSelection.getRangeAddress()
    return oArea.StartColumn + horizontalOffset, oArea.StartRow + verticalOffset, oArea.EndColumn + horizontalOffset, oArea.EndRow + verticalOffset
    
def reverse(aTuple):
    return tuple(t for t in reversed(aTuple))

# main function
def main():
       
    # get the Cell Range
    # use tuple unpacking of getSelectionAddresses returned tuple as parameter of getCellRangeByPosition
    oRangeSource = active_sheet.getCellRangeByPosition(*getSelectionAddresses())
    
    # get data from the Range of cells and store in a tuple
    oDataSource = oRangeSource.getDataArray()
    
    # print to console
    print(oDataSource)
    
    # create a new range of cells
    # from current selection using function getSelectionAddresses with offset
    # you can use the kwarg (keyword argument) horizontalOffset in parameter
    oRangeTarget = active_sheet.getCellRangeByPosition(*getSelectionAddresses(horizontalOffset = 3))
    
    # Reverse the data 
    oReversedSource = reverse(oDataSource)
    # Then set data for the target range using the 'reversed' data from the source range.
    oRangeTarget.setDataArray(oReversedSource)
    
