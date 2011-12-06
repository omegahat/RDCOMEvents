\name{createEventServerFromTemplate}
\alias{createEventServerFromTemplate}
\title{Create native DCOM server instance from meta-information.}
\description{
 This function merges user-supplied methods and a template
 for an event handler and creates an instance of the server
 using the R code as methods.
}
\usage{
createEventServerFromTemplate(methods, template, interfaceName)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{methods}{a named list of functions giving the methods that 
        are to be implemented in R by this server instance.
        Methods not specified are degenerate and are not called.}
  \item{template}{a template describing the interface for the "standard"
    or expected event handler.  This gives information about all the 
   methods, their identifiers and the types of the parameters and the 
   return values.
   See \code{\link{createCOMEventServerInfo}}.}
  \item{interfaceName}{the human-readable name of the event interface being implemented,
    e.g. AppEvents, WorkbookEvents, etc.  This is only used for providing meaningful
    error messages.}
}
\value{
  This returns the value from \code{\link{createCOMEventServer}},
  so an object containing a  reference to the low-level C++ object
  that represents the instance of the DCOM event handler.
}
\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{createCOMEventServer}}
 \code{mergeFunctions} which handles inserting the contracts for coercing the 
   arguments to the expected types.
 \code{\link{createCOMEventServerInfo}}
}
\examples{
}
\keyword{interface}
\concept{distributed computing}
\concept{event programming}
\concept{bi-directional interfaces}

