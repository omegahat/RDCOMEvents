\name{RDCOMNativeEventServer-class}
\docType{class}
\alias{RDCOMNativeEventServer-class}
\title{Event Server Instance class}
\description{
 This class is used to represent an instance of an R-implemented
 DCOM server. This is the object that references the C++-object
 that implements the server.
}
\section{Objects from the Class}{
 Instances of this class are typically created indirectly via
 \code{createCOMEventServer} and other higher-level functions
 such as \code{connectEventHandlers}.
}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"},
           inherited from \code{\link[RDCOMServer]{RDCOMServer-class}}}
  }
}
\section{Extends}{
Class \code{"RDCOMServer"}, directly.
}
\section{Methods}{
  \describe{
  }
}
\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{connectEventHandlers}}
 \code{\link{createCOMEventServer}}
 \code{\link{createCOMEventServerInfo}}
}
\examples{
}
\keyword{classes}
