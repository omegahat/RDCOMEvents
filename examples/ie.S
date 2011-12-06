# This doesn't seem to work - October 22, 2007
# The event handlers are not invoked.

library(RDCOMClient)
library(RDCOMEvents)
library(SWinTypeLibs)

app = COMCreate("InternetExplorer.Application")
app[['visible']] = TRUE
app$Navigate("http://www.google.com")
	
connectLinks = function(app)
{
doc = app$Document()
lib = LoadTypeLib(doc)
#lib = LoadTypeLib("c:/WINDOWS/system32/mshtml.tlb")
pts = getConnectionPoints(doc)
pt = pts[[2]]

iface = lib[[names(pts)[2]]]



s = createCOMEventServerInfo(iface, 
	    methods = list(StatusTextChange = function(...) {
	                                        cat("StatusTextChange", nargs(), "\n")
                                              },
	        	DownloadComplete = function(...) {
	                                        cat("DownloadComplete", "\n")
                                              }), complete = TRUE)
	
server = createCOMEventServer(s, direct = TRUE, verbose = TRUE)

cookie = connectConnectionPoint(pt, server)

list(info = s, server = server, cookie = cookie)
}

connectLinks(app)

#
	
# Now click the Images link in google
