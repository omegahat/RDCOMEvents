# Example

library(SWinTypeLibs)
library(RDCOMClient)
library(RDCOMEvents)
	
e = COMCreate("Excel.Application")
lib = LoadTypeLib(e)
	
 funcs = getFuncs(lib[["IAppEvents"]])

   # Template functions with parameters
funcDefs = createFunctions(funcs, createCompiledFunction, lib)
 
myFuncs = list(
      NewWorkbook = function(book) {  cat(book[["Name"]], "\n")  },
      SheetChange = function(sheet, target) {  print(sheet)
                                               print(target)
                                            },
      WorkbookBeforeSave = function(workbook, save, cancel) {
                              cat("Saving workbook\n")
                           },  

  
      WorkbookBeforeXmlImport = function(book, map, uri, is.refresh, cancel) {
                                   cat("importing xml\n")
                                }
     )


handlers =  mergeFunctions(myFuncs, funcDefs)
	

