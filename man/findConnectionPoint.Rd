\name{findConnectionPoint}
\alias{findConnectionPoint}
\alias{findConnectionPoint,COMIDispatch,character-method}
\alias{findConnectionPoint,COMIDispatch,ITypeInfo-method}
\title{Find specific connection point by class}
\description{
  To receive notification of events, on needs to first
  locate the relevant event source or Connection Point.
  A COM object can support multiple different event interfaces,
  each of which corresponds to a different connection point.
  This function queries the specified
  COM object (i.e. its \code{IConnectionPointContainer} object)
  for a particular connection point instance identified
  by the class of the event interface.
  This allows one to ask for a particular connection point
  rather than enumerating all of them via
  \code{\link{getConnectionPoints}}.
  It is also necessary when   \code{\link{getConnectionPoints}}
  is not implemented by the particular COM object (e.g. Active X controls).
}
\usage{
findConnectionPoint(obj, iid, expand = TRUE)
}
\arguments{
  \item{obj}{the COMIDispatch object in which to look for the connection
  point. This is the source of the events.}
  \item{iid}{the class identfier which specifies the particular event
    interface of interest.
    This can be a UUID (in string form and with the containing "{" and
    "}", or alternatively an \code{ITypeInfo} object obtained
    directly from the type library using the \code{SWinTypeLibs}
    package.
    The latter is currently used when converting the UUID from a string
    to a CLSID (via the Microsoft API routine \code{CLSIDFromString})
    fails.
  }
  \item{expand}{a logical value. If this is \code{TRUE}, the basic
    \code{\link{IConnectionPoint-class}} is turned into a
   \code{\link{IExpandedConnectionPoint-class}}   and this
   provides information about the UUID of the interface
   needed to implement the methods for this connection point
   and also the parent object associated with the connection point.
   This information can then be used to simplify the connection
   to the event source as the object becomes self-describing.
  }
}
\details{
 This calls \code{FindConnectionPoint} in the Windows API.
}
\value{
 An object of class \code{\link{IConnectionPoint-class}}
}

\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
  \code{\link{getConnectionPoints}}
  \code{\link{connectConnectionPoint}}
}
\examples{
library(RDCOMClient)
   xls <- COMCreate("Excel.Application")
   xls[["Visible"]] <- TRUE
   xls[["Workbooks"]]$Add()

   sh1 <- xls[["Sheets"]]$Item(1)
   oles <- sh1[["OLEObjects"]]
   oleButton <- oles$Add(ClassType = "Forms.CommandButton.1",
                Top = 10, Left = 10, Width = 144, Height = 30)

   acx <- oleButton[["Object"]]
   acx[["Caption"]] <- "Compute Means"

    # Now find the connection point, i.e. the event source
    # that we can register with to receive events on the button.
    # Do this by looking in the type MSForms library 
   library(SWinTypeLibs)
   formLib = LoadTypeLib("C:\\\\WINDOWS\\\\System32\\\\FM20.DLL")

\dontrun{
   pt = findConnectionPoint(acx, formLib[["CommandButtonEvents"]])

    # Now we have the event source, we create the handler.
    # See RDCOMServer.
   library(RDCOMServer)
   sinfo = createCOMEventServerInfo(formLib[["CommandButtonEvents"]],
                                     methods = list("Click"=function(...){cat("Hi from click\n")}))

   server = createCOMEventServer(sinfo)

   connectConnectionPoint(pt, server)
}

   xls$Quit()
}
\keyword{interface}

