#
# This file provides the facilities to deal with IConnectionPoint
# objects - obtaining them, connecting and disconnecting event handlers to them.
#


setClass("IConnectionPoint", contains = "IUnknown")
setClass("IExpandedConnectionPoint",
           representation(guid= "character",
                          source = "COMIDispatch"),
           contains = "IConnectionPoint")

getConnectionPoints = 
#
# Used to get some or all of the connection points
# ids is a character vector of UUIDs against which we can match.
#
function(obj, ids = NULL, expand = TRUE)
{
  points = .Call("R_getEnumConnectionPoints", obj, PACKAGE="RDCOMEvents")

  if(length(ids))
    points = points[as.character(ids)]
  else  {
     # Just return the unique IID values, i.e the names.
     # Can't easily perform unique() on the IConnectionPoint objects.
    points = points[match(unique(names(points)), names(points))]
  }

  if(expand) {
     ids = names(points)
     points = lapply(ids, function(id) {
	                    ExpandedConnectionPoint(points[[id]], id, obj)
                          })

     names(points) = ids
  }
  points
}

ExpandedConnectionPoint =
function(point, uuid, source, class = "IExpandedConnectionPoint")
{
   new(class,
        ref = point@ref,
        guid = uuid,
        source = source)
}

setGeneric("findConnectionPoint", function(obj, iid, expand = TRUE) standardGeneric("findConnectionPoint"))

setMethod("findConnectionPoint",
           c("COMIDispatch", "character"),
           function(obj, iid, expand = TRUE) {
              pt = .Call("R_FindConnectionPoint", obj, iid, PACKAGE="RDCOMEvents")
              if(expand)
                 pt = ExpandedConnectionPoint(pt, iid, obj)
	      pt
           })

setMethod("findConnectionPoint",
           c("COMIDispatch", "ITypeInfo"),
           function(obj, iid, expand = TRUE) {
             x = .Call("R_FindConnectionPoint", obj, iid@ref, PACKAGE="RDCOMEvents")
             if(is(x, "list"))
               x = x[[1]]
             if(expand) {
#XXX get the iid
                 x = ExpandedConnectionPoint(x, character(0), obj)
             }
             x
           })



connectConnectionPoint = Advise =
function(point, server)
{
  if(!is(point, "IConnectionPoint"))
    stop("point must be an IConnectionPoint object")
  
  .Call("R_AdviseConnectionPoint", point, server, PACKAGE="RDCOMEvents")
}

disconnectConnectionPoint = Unadvise =
function(point, cookie)
{
  if(!is(point, "IConnectionPoint"))
    stop("point must be an IConnectionPoint object")
  
  .Call("R_UnadviseConnectionPoint", point, cookie, PACKAGE="RDCOMEvents")
}




getUUIDFromConnectionPoint =
function(point)
{
  .Call("R_getConnectionInterface", point)
}
