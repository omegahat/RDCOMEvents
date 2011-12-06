\name{getConnectionPoints}
\alias{getConnectionPoints}
\title{Retrieve Event Source objects in a COM object}
\description{
 This gets the COM connection points
 or event source objects from a given COM object.
 Each of these  is an event source 
 for a  different interface or collection of event handler
 methods.
 One can connect an appropriate COM object as an event handler to one
 of these sources using \code{\link{connectConnectionPoint}}.
}
\usage{
getConnectionPoints(obj, ids = NULL, expand = TRUE)
}
\arguments{
  \item{obj}{the COM object in which we find the connection point instances.}
  \item{ids}{if specified, a character vector giving the stringified UUIDs of
   the connection points of interest. This allows one to only retrieve
   that particular value, although it is not implemented that way currently.}
  \item{expand}{a logical value. If this is \code{TRUE}, the basic
    \code{\link{IConnectionPoint-class}} is turned into a
   \code{\link{IExpandedConnectionPoint-class}}   and this
   provides information about the UUID of the interface
   needed to implement the methods for this connection point
   and also the parent object associated with the connection point.
   This information can then be used to simplify the connection
   to the event source as the object becomes self-describing.
  }
}
\value{
 A list with as many S \code{\link{IConnectionPoint-class}}
 or \code{\link{IExpandedConnectionPoint-class}} objects
 as there are in the underlying COM object.
 The names of the elements in the list are the 
 stringified UUIDs identifying the COM interface
 for the event handler.
}

\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang (duncan@wald.ucdavis.edu)}


\seealso{ 
 \code{\link{connectConnectionPoint}}
 \code{\link{disconnectConnectionPoint}}
}
\examples{
 e = COMCreate("Excel.Application")
 src = getConnectionPoints(e)[[1]]

 e$Quit()
 rm(list=c("e", "src"))
 gc()
}
\keyword{interface}


