library(RDCOMClient)
library(RDCOMServer)
library(SWinTypeLibs)

xls <- COMCreate("Excel.Application")
xls[["Visible"]] <- TRUE
xls[["Workbooks"]]$Add()
sh1 <- xls[["Sheets"]]$Item(1)

## generate and export random groups and data

N <- 200
grps <- sample(1:3, size = N, replace = TRUE)
x <- rnorm(N, mean = grps, sd = 0.5)
grps <- c("Low", "Med", "High")[grps]

at <- sh1$Range("A5")
exportVector(grps, at)
at <- at[["Next"]]
exportVector(x, at)

## add a ListBox ActiveX control and initialize

ole.lb <- addControl(sh1, "ListBox", 
            top = 10, left = 10, 
            width = 1.5 *72, height = 45)
lb <- ole.lb[["Object"]]
for(w in unique(grps)){
   cat(w, "\n")
   lb$AddItem(w)
}
lb$Activate()
      
# the follwing will be our event handler

cmpMedian <- 
function()
{
   Cols <- xls[["ActiveSheet"]][["UsedRange"]][["Columns"]]
   g <- importVector(Cols$Item("A"), header = FALSE)
   x <- importVector(Cols$Item("B"), header = FALSE)
   w <- lb[["Value"]]
   i <- which(g==w)
   m <- median(x[i], na.rm = TRUE)
   xls$InputBox(Prompt=paste("Group ",w, " (",
                             length(i)," observations)",sep=""), 
                Default = paste("median =", format(round(m, 3))))
}

ee <- setControlEvents(lb, "ListBox", Click = cmpMedian)

if(FALSE)
   deleteEventHandler(ee)

