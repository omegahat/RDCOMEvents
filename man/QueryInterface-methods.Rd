\name{QueryInterface-methods}
\docType{methods}
\alias{QueryInterface-methods}
\alias{QueryInterface,ANY,ANY-method}
\alias{QueryInterface,RDCOMEventServer,character-method}
\title{Method for RDCOMEventServer for querying if interface
       is supported by server. }
\description{
 This method provides an implementation for
 the RDCOMEventServer of the DCOM 
 \code{IUnknown} interface method to query 
 whether the server implements a specific interface.
 Such a query by a client to this R-implemented server
 results in a call to the \code{\link[RDCOMServer]{QueryInterface}}
 and this method.
 The method merely consults the vector GUIDs in the server
 and sees if the particular GUID is in that collection.

 This method could, and probably should, be moved
 to the \code{RDCOMServer} package if 
 \code{RDCOMEventServer} was made a proper class 
 that also inherited  from the COM server.
}
\section{Methods}{
\describe{

\item{obj = "ANY", guid = "ANY"}{ the generic!. This is 
   in the \code{RDCOMServer} package. }

\item{obj = "RDCOMEventServer", guid = "character"}{
     the method consults the GUIDs provided by the server
     that it implements and compares the GUID specified
     in the query to those. It returns a logical value
     indicating whether the given GUID is in the server's
     GUID collection.
  }
}}
\keyword{methods}
\concept{distributed computing}
\concept{DCOM}

