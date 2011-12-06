library(RDCOMClient)
library(RDCOMEvents)

app = getCOMInstance("Word.Application")
button = app$CommandBars()$Item(1)$Controls()$Item(1)

z = getConnectionPoints(button)	

if(TRUE) {
v = list()
for(j in 1:3) {
 try(
  for(i in 1:4)	 {
     cat("[ConnectionPointsTest] iteration", i, "of", j, "\n")
     z = getConnectionPoints(button)	
     cat("    got", length(z), "\n")
     v[[i]] = z
  }
 )
}
}