\name{IConnectionPoint-class}
\docType{class}
\alias{IConnectionPoint-class}
\alias{IExpandedConnectionPoint-class}
\title{S reference to a COM event source, aka Connection Point}
\description{
  The \code{IConnectionPoint} class is used to represent
  the objects with which one can register event handlers
  in R to DCOM event sources.
  Such objects are references to  COM objects that fire events
  to interested parties. We use this data type when registering
  an event handler with such a COM object via \code{\link{connectConnectionPoint}}

  The \code{IExpandedConnectionPoint} is a derived class that is 
  intended to contain additional information that makes the
  connection point self-describing for the purposes of use
  in other S functions.   It contains a reference
  to the DCOM object from which the connection point
  was retrieved. It also contains the UUID
  of the interface that an event handler must implement
  to register for events from that source.
}
\section{Objects from the Class}{
  One does not create \code{IConnectionPoint}
  objects manually. Instead, one must obtain
  them via the functions
  \code{\link{getConnectionPoints}}
  and \code{\link{findConnectionPoint}}.
}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr", from class "IUnknown"}
    This is the reference to the COM object which represents the event source.}
    \item{\code{source}:}{Object of class \code{"COMIDispatch"}.
     This is the object used to query the connection point.
     It is essentially the host of the connection point.
     This can be used to find the type library that provides
     descriptions about the interface that is expected for event
     sinks/handlers for this connection point.}
    \item{\code{guid}:}{Object of class \code{"character"}
    This is the identifier for the interface that the event handler must
    implement to listen for events from this source/connection point.
    One can use this information to find the definition of the interface
    within the type library associated with the source of the connection point,
    i.e. the value of the \code{source} slot.}
  }
}
\section{Extends}{
Class \code{"IUnknown"}, directly.
}
\section{Methods}{
No methods defined with class "IConnectionPoint" in the signature.
}

\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@wald.ucdavis.edu)}


\seealso{ 
\code{\link{connectConnectionPoint}}
\code{\link{getConnectionPoints}}
\code{\link{findConnectionPoint}}
\code{\link[RDCOMClient]{COMIDispatch-class}}
}

\examples{

}
\keyword{classes}
