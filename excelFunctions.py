def convertExcelColumn( column):
    match column: 
        case num if num in range(1,27):
            return chr(64+column)
        case num if num in  range(27,704):
            return chr(64+(int)((column-1)/26))+chr(65+((column-1)%26))





def getDefinedNameValue(workbook, definedName):
    dest = workbook.defined_names[definedName].destinations
    
    
    for title, coord in dest:
        cell = workbook[title][coord]
    
    return cell.value