library(SWinTypeLibs)
library(RDCOMClient)
library(RDCOMServer)

xls <- COMCreate("Excel.Application")
xls[["Visible"]] <- TRUE
xls[["Workbooks"]]$Add()
sh1 <- xls[["Sheets"]]$Item(1)

data(USArrests)
exportDataFrame(USArrests, sheet = sh1)

oleButton <- 
   addControl(sh1, "CommandButton",
              Top = 10, Left = 5*72, Width = 144, Height = 30)
acx <- oleButton[["Object"]]
acx[["Caption"]] <- "Compute Means"

computeMeans <- function()
{
   sh1 <- xls[["ActiveSheet"]]
   x <- importDataFrame(sheet = sh1)
   avgs <- colMeans(x)
   rws <- sh1[["UsedRange"]][["Rows"]]
   bottom <- rws$Item(rws[["Count"]])
   lbl.pos <- bottom$Offset(RowOffset = 1, ColumnOffset = 0)$Cells(1,1)
   lbl.pos[["Value"]] <- "Means:"
   avgs.pos <- lbl.pos[["Next"]]
   here <- avgs.pos
   RGB <- function(red, green, blue) red + green * 256.0 + blue * 256.0 * 256.0
   for(i in 1:length(avgs)){    ## bold font for all means
      font <- here[["Font"]]
      font[["Bold"]] <- TRUE
      font[["Color"]] <- RGB(255, 0, 0)
      here <- here[["Next"]]
   }
   exportVector(avgs, at = avgs.pos, byrow = TRUE)
   ## now set view to the row of Means...
   avgs.pos$Activate()
}

eh <- setControlEvents(acx, "CommandButton", Click = computeMeans)

if(FALSE){
   deleteEventHandler(eh)
}

