library(RDCOMClient)
library(RDCOMEvents)

word = getCOMInstance("Word.Application")
ctrl = word$CommandBars()$Item(1)$Controls()$Item(1)

ctr = 1
while(ctr < 100) {
  points = try( getConnectionPoints(ctrl))
  if(inherits(points, "try-error") || length(points) == 0)
    break
  ctr = ctr + 1
}

cat(" Number of iterations", ctr, "\n")
