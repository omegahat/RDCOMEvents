\name{connectEventHandlers}
\alias{connectEventHandlers}
\alias{connectEventHandlers,IConnectionPoint,list,ANY-method}
\alias{connectEventHandlers,COMIDispatch,RDCOMNativeEventServer,ANY-method}
\alias{connectEventHandlers,COMIDispatch,list,character-method}
\title{Simple facility for registering functions as event handlers for a connection point.}
\description{
 This function  is just a simplified version of a sequence of
 steps used to create a COM event server from a collection of
 methods and an interface definition and then connect this
 server to an event source/connection point.
 This function makes it a single action.
 It leverages the \code{\link{IExpandedConnectionPoint-class}}
 class to provide a self-describing connection point.
}
\usage{
connectEventHandlers(eventSource, methods, 
                     lib = LoadTypeLib(eventSource@source),
                     iface = lib[[eventSource@guid]])
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{eventSource}{the connection point. This is expected to be of class 
   \code{IExpandedConnectionPoint}}
  \item{methods}{a named list of functions.  These are the functions
    that implement the methods of the event handler interface expected for this
    connection point.  These are passed to 
    \code{\link{createCOMEventServerInfo}}.}
  \item{lib}{the type library that contains the definition of the interface 
    for this connection point's events.  This is computed from the 
    \code{source} field of the \code{eventSource} argument.
    It is a parameter to this function to allow those who have already
    loaded the library to pass it directly.
  }
  \item{iface}{this ITypeInfo object that provides the definition of the
    interface for an event sink/handler for this connection  point.
    If this is specified, the \code{lib} is not used.
    Therefore, if this has been pre-computed, it can be provided as
    an argument to this function call. Otherwise, it is computed
    using the UUID in the \code{\link{IExpandedConnectionPoint-class}}
    object given in \code{eventSource}.
  }
}
\details{

}
\value{
 A list with 2 elements:
 \item{server}{the COM server instance that was registered with the event source.}
 \item{cookie}{the identifier or cookie returned from  
   \code{\link{connectConnectionPoint}} that can be used
   to unregister/disconnect the event handler.}
}


\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{findConnectionPoint}}
 \code{\link{getConnectionPoints}}

 \code{\link{createCOMEventServerInfo}}
 \code{\link{createCOMEventServer}}

 \code{\link{connectConnectionPoint}}
 \code{\link{Unadvise}}
}
\examples{

\dontrun{
  library(RDCOMClient)
  library(SWinTypeLibs)	 
  e = getCOMInstance("Excel.Application")
  point = getConnectionPoints(e)[[1]]

  v = connectEventHandlers(point,
                            list(SheetChange = function(Sh, Target) {
                                                  cat("Sheet has been changed\n")
                                               }))
}
}
\keyword{programming}
\keyword{interface}
\concept{distributed computing}
\concept{event-driven computing}
\concept{DCOM}


