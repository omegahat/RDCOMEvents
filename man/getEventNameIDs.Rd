\name{getEventNameIDs}
\alias{getEventNameIDs}
\title{Retrieve the name-identifier for COM event methods}
\description{
  This computes the names and the associated identifiers
  for a collection of methods in a COM event interface.
  When an event is broadcast to interested recipients,
  the identifier is used to indicate the particular
  event. This mapping of name-identifiers is used
  to identify the actual name of the event so that
  we can match it to a function in S.
}
\usage{
getEventNameIDs(info)
}
\arguments{
  \item{info}{the \code{ITypeInfo} object that describes
    the interface of interest. This is typically retrieved
    using the \code{SWinTypeLibs} package.}
}
\value{
  A named integer vector giving the integer values corresponding
  to the different event  types and the associated names.
}

\references{
 \url{http://www.omegahat.org/RDCOMServer}
 \url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}



\seealso{
  \code{\link{createCOMEventServer}}
  \code{\link{getEventNameIDs}}
}
\examples{
 library(RDCOMClient)
 library(SWinTypeLibs)
 e = getCOMInstance("Excel.Application")
 lib = LoadTypeLib(e)
 e$Quit()

 library(RDCOMServer)
 i = lib[["IWorkbookEvents"]]
 RDCOMEvents:::getEventNameIDs(i)
}
\keyword{interface}

