library(SWinTypeLibs)
library(RDCOMClient)

e = COMCreate("Excel.Application")
pts = getConnectionPoints(e)

lib = LoadTypeLib("C:\\Program Files\\Microsoft Office\\Office\\EXCEL9.OLB")

