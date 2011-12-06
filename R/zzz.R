.First.lib =
function(lib, pkg) 
{
 library.dynam("RDCOMEvents", pkg, lib)
 library(RDCOMClient)
}
