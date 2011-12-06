\name{createCOMEventServerInfo}
\alias{createCOMEventServerInfo}
\alias{createCOMServerInfo}
\title{Create default handlers for COM event interface}
\description{
 This generates a collection of template (degenerate) functions 
 that implement the specified interface.
 This is used to generate the definition of a full server object
 for a particular DCOM interface so that methods can be overridden
 with those given in \code{methods}.
 The resulting object can be used to instantiate the 
 actual server object via \code{\link{createCOMEventServer}}.
}
\usage{
createCOMEventServerInfo(info, eventIDInfo = info, 
                         methods = NULL, .lib = NULL, 
                         complete = is.null(methods), 
                         name = "", help = "", 
                         def = new("COMEventServerInfo"))
}
\arguments{
  \item{info}{the \code{ITypeInfo} object in which to find the
    functions.
    This can be the name of the \code{ITypeInfo} element within the
    type library if \code{.lib} is specified.}
  \item{eventIDInfo}{the \code{ITypeInfo} object in which to lookup 
    the names and identifiers of the event types.
    This potentially can be different from \code{info}
    in that one might refer to the IDispatch interface
    and the other to the regular interface description.
    This can be the name of the \code{ITypeInfo} element within the
    type library if \code{.lib} is specified.}
  \item{\dots}{}
  \item{methods}{a named list of functions that are to be used as event
    handlers.
     These are used to replace the
    default or template functions created by reading the type
    information.
   }
 \item{.lib}{if \code{info} or \code{eventIDInfo} are given as strings,
   this should be the \code{ITypeLibrary} object  in which to find these
   elements.}
 \item{complete}{a logical value. If \code{TRUE}, then we create a list
   of functions for all of the methods for the specified interface.
   Otherwise, we use only the methods given by \code{\dots}.
   One typically uses \code{TRUE} to find out about the different
   methods and their arguments so that one can provide implementations
   in the future. In other words, this is more of a development tool.
  }
 \item{name}{a human-readable name for the COM class.
   This is the value assigned to the \code{name} slot in the
  \code{def} object.}
 \item{help}{a string giving a short description of the server.
   This might be used in implementing s}
 \item{def}{the object which is to be filled in by this function.
     This argument allows the caller to specify an instance of a
     derived (or compatible) class from \code{\link{COMEventServerInfo-class}}.
   }

}
\value{
  A list containing two elements
  \item{methods}{the collection of named functions created by merging
    the caller-specified functions (via \dots) into the template/default
    methods constructed from the event interface.
  }
  \item{ids}{the mapping from the event identifiers to names given as an
  integer vector and computed via \code{\link{getEventNameIDs}}}
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

}
\keyword{interface}

