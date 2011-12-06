\name{createCOMEventServer}
\alias{createCOMEventServer}
\alias{createCOMEventServer,list,numeric-method}
\alias{createCOMEventServer,COMEventServerInfo,ANY-method}
\alias{createCOMEventServer,COMIDispatch,list-method}
\title{Create S-level COM event dispatch handler.}
\description{
  Create an instance of an S-level COM server
  that handles dispatching events to S functions specified by the collection
  of named functions.
  The dispatcher maps the event types to the functions using the mapping
  of identifiers to names given by \code{ids}.
  These identifiers are computed from the ITypeInfo
  in the type library for the event source's interface.
  This relates the names of the methods to unique identifiers
  that were determined when the interface was defined.
  The collection of functions provide methods for the
  events of interest and are invoked when an event is received
  by the C-level COM server.
}
\usage{
createCOMEventServer(methods, ids, guids = character(0),
                       direct = TRUE, verbose = FALSE)
}
\arguments{
  \item{methods}{a named list of S functions that are considered
    handlers for the different events that the source can generate.
   }
   \item{ids}{a mapping between the event names and their internal
     identifiers that are used when the event is received by this
     low-level event handler.
   }
   \item{guids}{a character vector giving the UUIDs that this server
    implements. These are used when a DCOM client queries the 
    supported interfaces using the basic QueryInterface of 
    \code{IUnknown} interface.}
   \item{direct}{a logical value which, if \code{TRUE} causes the
     low-level C COM server to be created; if \code{FALSE}, just
     the S-level COM dispatch mechanism is returned which can be passed
     to the \code{R_RCOMSObject} C routine in a future call.
   }
   \item{verbose}{a logical value which if \code{TRUE} indicates that
     events with no handler method should be announced on the console.
     This is useful for debugging.}
}
\details{

}
\value{
  If \code{direct} is \code{FALSE}, 
  the return value is a list giving 
  the \code{Invoke} function that handles the dispatching
  of the COM calls to the S functions, and also a degenerate
  or \code{NULL} function for mapping names to identifiers.
  This is not needed as the event and argument names are already mapped
  by the definition of the event handler interface.

  If \code{direct} is \code{TRUE}, this function returns
  an S external pointer object that contains a reference
  to the underlying C++ COM object that interacts with DCOM
  and the event source.
}

\references{
 \url{http://www.omegahat.org/RDCOMServer}
 \url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
  \code{\link{createCOMEventServerInfo}}
  \code{\link{getEventNameIDs}}
}  

\examples{

}
\keyword{interface}

