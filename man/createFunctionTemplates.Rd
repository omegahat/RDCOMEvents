\name{createFunctionTemplates}
\alias{createFunctionTemplates}
\title{Generate stub/template functions for event handler methods}
\description{
 This function is used in the generation of R code from 
 information from a type library to provide a "compiled"
 interface to a collection of DCOM interfaces.
 This generation is implemented as part of the 
 \code{\link[SWinTypeLibs]{generateInterface}}
 in the \code{SWinTypeLibs} package.
 This function generates the basic definitions of the
 R functions provided in the given DCOM interface's methods.
 Depending on the choice of function  for the argument \code{op},
 information about the coercions of the parameters to R types
 and the return value is also encoded in the function.
}
\usage{
createFunctionTemplates(funcs, op = createCompiledFunction, ...)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{funcs}{the descriptions of the methods in the DCOM interface.
    See \code{\link[SWinTypeLibs]{getFuncs}}.}
  \item{op}{the function that is used to generate the
    stubs for each method description.  This is the heart of the 
    function.  The built-in choices are 
    \code{\link{createFunction}} and \code{\link{createCompiledFunction}}.
   The latter is greatly preferred as it includes information about the types
   of the parameters and return value.}
  \item{\dots}{additional arguments that are passed to the \code{op} function
  each time it is called, i.e. for each method definition.}
}
\value{
 A list of values returned from calling the \code{op} function for each 
 method.  For the two built-in functions, these are simply functions
 with special attributes.
}
\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{createCompiledFunction}}
 \code{\link{createFunction}}
 \code{\link[SWinTypeLibs]{generateInterface}}
}
\examples{

}
\keyword{programming}
\keyword{interface}
\concept{meta-programming}

