\name{COMEventServerInfo-class}
\docType{class}
\alias{COMEventServerInfo-class}
\title{S-level description of a COM event handler interface}
\description{
  Objects of this class contain information used to describe
  the different methods a COM event handler might implement for
  a particular event interface.
  This meta-information is combined with user-specified functions
  to implement a particular event handler.
}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{new("COMEventServerInfo", ...)}.

}
\section{Slots}{
  \describe{
    \item{\code{methods}:}{Object of class \code{"list"} a named list of
    functions that are used as the methods of the event handler. The
    names should correspond to the names in \code{ids}.}
  \item{\code{ids}:}{Object of class \code{"numeric"}, 
    inherited from \code{\link[RDCOMServer]{COMInterfaceServerInfo-class}}
  }
    \item{\code{classId}:}{inherited from \code{\link[RDCOMServer]{SCOMClass-class}}}
    \item{\code{guids}:}{inherited from \code{\link[RDCOMServer]{SCOMClass-class}}}
    \item{\code{help}:}{inherited from \code{\link[RDCOMServer]{SCOMClass-class}}}
    \item{\code{name}:}{inherited from \code{\link[RDCOMServer]{SCOMClass-class}}}
    \item{\code{progId}:}{inherited from \code{\link[RDCOMServer]{SCOMClass-class}}}
    \item{\code{version}:}{inherited from \code{\link[RDCOMServer]{SCOMClass-class}}}
  }
}

\section{Methods}{
  \describe{
    \item{createCOMEventServer}{\code{signature(methods =
	"COMEventServerInfo", ids = "ANY")}:
     constructor function for creating the description of an event
     handler from information in a type library.}
  }
}
\references{
 \url{http://www.omegahat.org/RDCOMServer}
 \url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
  \code{\link[RDCOMClient]{connectConnectionPoint}}
  \code{\link[RDCOMClient]{getConnectionPoints}}  
}

\examples{

}
\keyword{classes}
