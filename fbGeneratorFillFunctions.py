
from datetime import datetime

from sys import argv
import re
from datetime import datetime
from openpyxl import load_workbook
from excelFunctions import getDefinedNameValue
import re


def fillTitleFunction (workbook, unitName, newFileName ):



    with open(newFileName,'a') as new_file:
            
            fbName = getDefinedNameValue(workbook,"fbName")
            title = getDefinedNameValue(workbook,"title")
            author = getDefinedNameValue(workbook,"author")
            version = getDefinedNameValue(workbook,"version")
            
            if  re.search("()?&Unit&()?", fbName):
                fbName = fbName.replace("&Unit&",str(unitName))

            #copy other lines 
            new_file.write("FUNCTION BLOCK \"" + fbName + "\""+"\n")
            new_file.write("TITLE : "+title+"\n")    
            new_file.write("AUTHOR : "+author+"\n")          
            new_file.write("VERSION : "+version+"\n")


def fillStatFunction (dic, newFileName ):



    with open(newFileName,'a') as new_file:
            
        new_file.write("\n")
        new_file.write("VAR\n")

        for key in dic.keys():

            [fb ,par] = dic[key]
            #copy other lines 
            new_file.write(key+"  : ")
            new_file.write("\""+ fb+"\";\n")  

        new_file.write("END_VAR\n")
        

def fillNetworkFunction (workbook,functionName, dataName, unitName, tagname, par, newFileName ):

    functionSheet = workbook[functionName]                      
    dataSheet = workbook[dataName] 



    with open(newFileName,'a') as new_file:

        new_file.write("\n")

        for row in functionSheet.values:
            
            #get a tagname when there is a new title 
               

            for value in row:

                newLine = value
                for column in range (1,dataSheet.max_column+1):
                    
                    match column:

                        case 1: 
                            checkedWord = "&Unit&"
                            replacedWord = str(unitName)
                        
                        case 2: 
                            checkedWord = "&tagname&"
                            replacedWord = str(tagname)

                        case num if num in range(4,99): 
                            checkedWord = "&par"+str(column-3)+"&" 
                            replacedWord = str(par[column-4])
                                    
                    

                    if  re.search(f"()?{checkedWord}()?", value):
                        
                        newLine = value.replace(checkedWord,replacedWord)
                        

            #copy other lines 
            new_file.write(newLine)
            new_file.write("\n")

        new_file.write("\n") 
            

                      
            



