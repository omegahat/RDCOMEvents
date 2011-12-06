CXX=c:/MinGW/bin/g++
ifndef CXX
endif

 CC=c:/MinGW/bin/gcc
ifndef CC
endif

ifndef RHOME
 RHOME=d:/duncan/Projects/R/R-check
endif

-include local.config
ifndef RDCOMSERVER_SRC
 RDCOMSERVER_SRC=.
endif

ifndef RDCOMCLIENT_SRC
 RDCOMCLIENT_SRC=.
endif

DEFS=RDCOMEvents.def
DEFINES=-D_GNU_ -DNO_PYCOM_IPROVIDECLASSINFO -DUSE_R
# -DREGISTER_COM_OBJECTS_WITH_S  
#    for registering the COM objects into an S table.
# -DANNOUNCE_COM_CALLS
#    for writing diagnostics to the terminal about the different COM calls
#    as they happen.

CXXFLAGS=-g -I$(RHOME)/src/include $(DEFINES) -Wno-deprecated -I$(RDCOMSERVER_SRC)
CFLAGS=-g -I. -I$(RHOME)/src/include $(DEFINES) -I$(RHOME)/src/gnuwin32 -I$(RDCOMSERVER_SRC)

DLLTOOL=dlltool
DLLWRAP=dllwrap

DLLWRAP_FLAGS = --driver-name $(CXX) --dlltool-name $(DLLTOOL) --as $(AS) 
LDLIBS=-L$(RHOME)/src/gnuwin32 -lR -lole32 -loleaut32 -luuid



OBJS=events.o $(RDCOMCLIENT_SRC)/connect.o $(RDCOMSERVER_SRC)/RUtils.o $(RDCOMSERVER_SRC)/COMError.o $(RDCOMSERVER_SRC)/converters.o


RDCOMEvents.dll: $(OBJS) 
	$(DLLWRAP) $(DLLWRAP_FLAGS) -o $@ $(OBJS) \
	  --def $(DEFS) $(LDFLAGS) $(LDLIBS)	

connect.o: connect.cpp StdAfx.h

clean: 
	rm *.o
