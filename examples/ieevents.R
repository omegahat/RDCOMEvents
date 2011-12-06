library(RDCOMClient)
library(RDCOMEvents)
library(SWinTypeLibs)

app = COMCreate("InternetExplorer.Application")
app[['visible']] = TRUE
app$Navigate("http://www.google.com")


connectLinks = function(app, link = "images")
{
  links = app$Document()[["url"]]
  
   # Find the Images link.
  links = app$Document()$links()
  hrefs = sapply(1:links[["length"]], function(i) links$Item(i)[["href"]])
  index = grep(link, hrefs)
  
  link1 = app$Document()$links(index)  # Note that their arrays here start at zero, unlike they do in office
  lib = LoadTypeLib("c:/WINDOWS/system32/mshtml.tlb")
  pts = getConnectionPoints(link1)
  pt = pts[[2]]
  iface = lib[[names(pts)[2]]] # Don't ask how I came up with 2, there are 
                               # five connection points all together: 1, 4, 
		   	       # and 5 aren't found in the type library, but 2 and 3 are

  s = createCOMEventServerInfo(iface, methods = list(onclick = function(...) {cat('YOU CLICKED ME! HOORAY! \n')}), complete = TRUE)
  server = createCOMEventServer(s, direct = TRUE, verbose = TRUE)

  cookie = connectConnectionPoint(pt, server)

  list(info = s, server = server, cookie = cookie, connections = pts, link = link1)
}

# Now click the Images link in google
