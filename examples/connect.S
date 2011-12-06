  library(RDCOMClient)
  library(RDCOMEvents)
  library(SWinTypeLibs)	 

#  e = getCOMInstance("Excel.Application")
  e = COMCreate("Excel.Application")
  book = e$Workbooks()$Add()
	
  point = getConnectionPoints(book)[[1]]

  v = connectEventHandlers(point,
                            list(SheetChange = function(Sh, Target) {
                                                  cat("Sheet has been changed\n")
                                               },
		                 SheetActivate = function(Sh) {
                                                  cat("SheetActivate\n")
                                               }))

  e[["Visible"]] = TRUE		