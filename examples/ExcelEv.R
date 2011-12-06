# See ExcelEvSimple.R for a simpler way to do this 
# that avoids dealing with the event server info and 
# creating the server directly.

library(RDCOMClient)
library(RDCOMEvents)
library(SWinTypeLibs)

e = COMCreate("Excel.Application")
book = e[["Workbooks"]]$Add()

connections = getConnectionPoints(book)

#LibFile = "C:/Microsoft Office/OFFICE11/Excel.exe"
#lib = LoadTypeLib(LibFile)

lib = LoadTypeLib(e)
iface = lib[[names(connections)]]
s = createCOMEventServerInfo(iface, complete = TRUE)

s@methods$SheetActivate = function(...) { cat("<SheetActivate>\n"); print(list(...)); cat("</SheetActivate>\n")}
s@methods$Activate = function(...) { cat("<WorkbookActivate>\n"); print(list(...)); cat("</WorkbookActivate>\n")}
s@methods$NewSheet = function(...) { cat("<NewSheet>\n"); print(list(...)); cat("</NewSheet>\n")}
s@methods$SheetChange = function(...) { cat("<SheetChange>\n"); print(list(...)); cat("</SheetChange>\n")}

s@methods$WindowResize = function(...) { cat("<WindowResize>\n"); print(list(...)); cat("</WindowResize>\n")}

e[["Visible"]] = TRUE

server = createCOMEventServer(s@methods, s@ids, direct = TRUE)
connectConnectionPoint(connections[[1]], server)




