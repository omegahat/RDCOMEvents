\name{createCompiledFunction}
\alias{createCompiledFunction}
\alias{createFunction}
\title{Generators for template functions for DCOM methods}
\description{
 These functions are used to create template R functions
 from descriptions of DCOM methods.  Using
 \code{createCompiledFunction}, these generated functions
 have the appropriate number of parameters, 
 expressions to coerce the arguments to the known types,
 and have the  return information associated with attributes.
 The simpler \code{createFunction} creates the simpler version
 that does not know about the types of the arguments or the
 return type.
}
\usage{
createCompiledFunction(func, lib)
}
\arguments{
  \item{func}{the }
  \item{lib}{the ITypeLib that }
}

\value{
An R function.
}

\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{createCOMServerInfo}}
}
\examples{

}
\keyword{programming}
\keyword{interface}

