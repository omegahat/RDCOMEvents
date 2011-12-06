# Load the  relevant libraries.
library(RDCOMClient)
library(RDCOMEvents)
library(SWinTypeLibs)

# Create an instance of Excel and add a new workbook.
e = COMCreate("Excel.Application")
book = e[["Workbooks"]]$Add()

# Get the one and only connection point for this workbook.
connections = getConnectionPoints(book)

# Create o
methods = list()
methods$SheetActivate = function(...) { cat("<SheetActivate>\n"); print(list(...)); cat("</SheetActivate>\n")}
methods$Activate = function(...) { cat("<WorkbookActivate>\n"); print(list(...)); cat("</WorkbookActivate>\n")}
methods$NewSheet = function(...) { cat("<NewSheet>\n"); print(list(...)); cat("</NewSheet>\n")}
methods$SheetChange = function(...) { cat("<SheetChange>\n"); print(list(...)); cat("</SheetChange>\n")}
methods$WindowResize = function(...) { cat("<WindowResize>\n"); print(list(...)); cat("</WindowResize>\n")}

# This version is simpler than the one in ExcelEv.R using a convenience
# function to create the event handler object and connect it to the
# first connection point.
handlerInfo = connectEventHandlers(connections[[1]], methods)

e[["Visible"]] = TRUE





