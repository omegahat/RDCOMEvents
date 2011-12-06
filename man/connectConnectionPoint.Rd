\name{connectConnectionPoint}
\alias{connectConnectionPoint}
\alias{Advise}
\alias{disconnectConnectionPoint}
\alias{Unadvise}
\title{Register/Unregister COM Event Handler}
\description{
These functions associate and disassociate
a COM server in S to an event source, or connection
point in COM-speak.
When connected, any events that occur in that
event source will cause the associated
method in the COM server's interface to be invoked.
These COM servers can be implemented entirely within
R using the \code{RDCOMServer} package.

When one wants to stop receiving notification of events, one can
disconnect from the event source.  This is done by passing back the
handle created when connecting to the event source.
}
\usage{
connectConnectionPoint(point, server)
}
\arguments{
  \item{point}{ the event source from which the event notification comes. 
    This is obtained using \code{\link{getConnectionPoints}}.
   }
  \item{server}{a COM server object, typically created 
   using \code{\link[RDCOMServer]{createCOMObject}}
   or \code{\link[RDCOMServer]{createCOMEventServer}}.}
}
\details{
 These use the C++-level  COM mechanism to communicate
 with the COM connection point.
}
\value{
 \code{connectConnectionPoint} returns
 the handle that identifies the particular connection.
 This is an opaque value and should never be used. 

 Disconnecting either returns \code{NULL} or 
 raises an error.
}
\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang (duncan@wald.ucdavis.edu)}

\seealso{ 
 \code{\link[RDCOMServer]{createCOMObject}}
 \code{\link[RDCOMServer]{createCOMEventServer}}
 \code{\link[SWinTypeLibs]{LoadTypeLibrary}}

 \url{event.html} 
}
\examples{
\dontrun{
 e = COMCreate("Excel.Application")
 con = getConnectionPoints(e)[[1]]

 library(RDCOMServer)
 createCOMEventServer(excelLib[["IAppEvents"]], )

 cookie = connectConnectionPoint(con, server)
 
   # Later, stop processing events.
 disconnectConnectionPoint(con, cookie) 
}
}
\keyword{interface}

