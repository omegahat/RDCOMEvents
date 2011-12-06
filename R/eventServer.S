#
#
# This file may be incorporated into RDCOMEvents.
# And indeed, it seems to have been!!



# This class holds the definition of the implementation of a given
# interface.  We provide the methods we want to override and
# the entire collection is stored.
setClass("COMEventServerInfo",
           representation(methods="list"), 
	   contains = "COMInterfaceServerInfo")

IUnknownMethodNames = c("QueryInterface", "AddRef", "Release")
IDispatchMethodNames = c("GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke")

# A constructor for COMEventServerInfo objects.
# This is similar in spirit to SCOMFunctionClass(), SCOMIDispatch()
# 
createCOMEventServerInfo = createCOMServerInfo  =
function(info, eventIDInfo = info, 
          methods = NULL, .lib = NULL, complete = is.null(methods), 
          name = "", help = "", def = new("COMEventServerInfo"))
{
 library(RDCOMServer)

 def = .initSCOMClass(def, name = name, help = help)
 def@guids = info@guid
	
 if(is.character(info) && !is.null(.lib))
   info = .lib[[info]]

 guids = info@guid

 if(is.character(eventIDInfo) && !is.null(.lib))
   eventIDInfo = .lib[[eventIDInfo]] 

 actualMethods = methods

  # If we have to complete the list of event methods
  # then find them and construct them.
 if(complete) {
  funcs = getFuncs(info, .lib)
  # Throw away the IDispatch ones.

  if(names(info@type) == "dispatch") {
    idx = match(c(IUnknownMethodNames, IDispatchMethodNames), names(funcs))
    if(any(!is.na(idx))) {
      funcs = funcs[- idx [ !is.na(idx) ]] 
    }
  }

  methods = createFunctionTemplates(funcs)

  if(length(actualMethods))
    methods[names(actualMethods)] = actualMethods
 } else
    methods = actualMethods

 ids = getEventNameIDs(eventIDInfo, .lib)
 if(length(ids) && !complete) {
     # remove the ids that do not have a method.
     #XXX We probably want to hold onto the return values however.
   which = match(names(methods), names(ids), 0)
   if(any(which == 0))
     warning("Handlers without corresponding event identifiers:", paste(names(methods)[which == 0], collapse=""))
   ids = ids[which]
 }

 if(length(ids) == 0)
    ids = numeric()
	
 def@methods = methods
 def@ids = ids
 def@guids = guids

 return(def)
} 

setGeneric("createCOMEventServer", 
           function(methods, ids, guids = character(0), direct=TRUE, verbose =FALSE)
            standardGeneric("createCOMEventServer"))

setMethod("createCOMEventServer",
           c("COMEventServerInfo"),
           function(methods, ids, guids = character(0), direct = TRUE, verbose = FALSE) {
             createCOMEventServer(methods@methods, methods@ids, methods@guids, direct, verbose)
           })


setMethod("createCOMObject", 
          "COMEventServerInfo",
          function(def) {
              createCOMEventServer(def)
          })

setMethod("createCOMEventServer",
           c("list", "numeric"),
function(methods, ids, guids = character(0), direct = TRUE, verbose = FALSE)
{  
 invoke = 
   function(id, method, args, namedArgs) {

      funcName = names(ids)[which(ids == id)]

      if(length(funcName) && funcName %in% names(methods)) {
        eval(as.call(c(methods[[funcName]], rev(args))), env = globalenv())
      } else if(verbose) {
	cat("Ignoring event number", id, "(", class(id), ")", "\n")
      }
   }

 s = list(Invoke = invoke, GetNamesOfIDs = NULL, guids = guids)
 class(s) =  "RDCOMEventServer"
# s = new("RDCOMEventServer", invoke = invoke, guids = guids)

 if(direct) {
   library(RDCOMServer)
   serverPtr = .Call("R_RCOMSObject", s, PACKAGE="RDCOMServer")
   s = new("RDCOMNativeEventServer", ref = serverPtr)
 }
 
 s
})

#XXXX
if(FALSE) {
setClass("RDCOMEventServer",
          representation(Invoke = "function",
                         GetNamesOfIds = "",
                         guids = "character"))
} else {
#XXX
 setOldClass("RDCOMEventServer")
}

# A pointer to a C++ object.
setClass("RDCOMNativeEventServer", representation("RDCOMServer"))

setMethod("QueryInterface",  c("RDCOMEventServer", "character"),
          function(obj, guid) {
            guid %in% obj$guids
          })


getEventNameIDs =
function(info, lib = getTypeLib(info))
{
  tmp = names(getFuncs(info, lib))
#  if(names(info@type) == "dispatch")
#   tmp = tmp[-c(1:7)]
  el = sapply(tmp, getNameIDs, info)
  names(el) = tmp
  el
}


setGeneric("connectEventHandlers", 
             function(eventSource, methods, lib = LoadTypeLib(eventSource@source), 
	                iface = lib[[eventSource@guid]])
                 standardGeneric("connectEventHandlers")
             )


setMethod("connectEventHandlers", c("IConnectionPoint", "list"),
#
# This is intended to be a convenient, high-level function
# that hides the details of creating a COMEventServerInfo
# then a server and then connecting to the event source.
# This function will collapse all of those actions into a
# single action.
#
#  eventSource is 
#
#
function(eventSource, methods, 
          lib = LoadTypeLib(eventSource@source), 
          iface = lib[[eventSource@guid]])
{
  serverInfo = createCOMEventServerInfo(iface, methods = methods, complete = TRUE)
  server = createCOMEventServer(serverInfo, direct = TRUE)

  cookie = connectConnectionPoint(eventSource, server)

  return(list(cookie = cookie, server = server, serverInfo = serverInfo))
}
)



# e.g
#    connectEventHandler(COMCreate("Excel.Application"),  AppEventsHandler(list(... )))
# See codeGeneration.xml
setMethod("connectEventHandlers", 
           c("COMIDispatch", "RDCOMNativeEventServer"),
	   function(eventSource, methods, lib = LoadTypeLib(eventSource@source), 
                             iface = lib[[eventSource@guid]]) {

                guid =  getServerDetails(methods)$guids
                point = findConnectionPoint(eventSource, guid[1])
                cookie = connectConnectionPoint(point, methods)

		list(cookie = cookie, server = methods, point = point)
           })

setMethod("connectEventHandlers", 
           c("COMIDispatch", "list", "character"),
	   function(eventSource, methods, lib = LoadTypeLib(eventSource@source), 
                             iface = lib[[eventSource@guid]]) {

                # if lib is a human-readable name, see if there is a function
                # named  <lib>Handler
                if(exists(paste(lib, "Handler", sep = ""), mode = "function")) {
                  f = get(paste(lib, "Handler", sep = ""), mode = "function")
                  methods = f(methods)
                } else {
#XXX it is a UUID so we have to find the event handler from that.
                }

                connectEventHandlers(eventSource, methods)
           })


createEventServerFromTemplate = 
#
# This is given a template generated by a call to createCOMEventServerInfo()
#
#
# The error checking could and probably should be merged into 
# createCOMEventServer()
#
function(methods, template, interfaceName)
{
  ok = match(names(methods), names(template@methods))
  if(any(is.na(ok))) {
    stop("The interface ", interfaceName, " has no methods named ", 
           paste(names(methods)[is.na(ok)], sep = ", "))
  }

  methods = mergeFunctions(methods, template@methods)
  template@methods[names(methods)] = methods
print(names(methods))
  server = createCOMEventServer(template, direct = TRUE)

  server
}
	
