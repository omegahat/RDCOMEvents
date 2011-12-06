include VersionInfo

PKG_FILES=$(shell find R examples src man -not -path '*CVS*' -not -name '*~') DESCRIPTION

PACKAGE_DIR=$(RHOME)/src/library
PACKAGE_DIR=/tmp
INSTALL_DIR=${PACKAGE_DIR}/$(PKG_NAME)

include Install/GNUmakefile.admin

RDCOMSERVER_SRC=../RDCOMServer/src
RDCOMSERVER_SRC_FILES=RUtils.c RUtils.h converters.cpp converters.h RCOMObject.h COMError.cpp

RDCOMCLIENT_SRC=../RDCOMClient/src
RDCOMCLIENT_SRC_FILES=connect.cpp StdAfx.h

C_SOURCE=$(wildcard src/*.[ch] src/*.cpp) $(RDCOMSERVER_SRC_FILES:%=$(RDCOMSERVER_SRC)/%) $(RDCOMCLIENT_SRC_FILES:%=$(RDCOMCLIENT_SRC)/%) src/Makefile  src/RDCOMEvents.def

R_SOURCE=$(wildcard R/*.[RS])
MAN_FILES=$(wildcard man/*.Rd)
INSTALL_DIRS=src man R

DOC_FILES=event.html event.xml
DOCS=$(DOC_FILES:%=inst/doc/%)

DESCRIPTION: DESCRIPTION.in Install/configureInstall.in VersionInfo

package: DESCRIPTION $(DOCS)
#	@if test -z "${RHOME}" ; then echo "You must specify RHOME" ; exit 1 ; fi
	if test -d $(INSTALL_DIR) ; then rm -fr $(INSTALL_DIR) ; fi
	mkdir $(INSTALL_DIR)
	cp NAMESPACE DESCRIPTION $(INSTALL_DIR)
	for i in $(INSTALL_DIRS) ; do \
	   mkdir $(INSTALL_DIR)/$$i ; \
	done
	cp -r $(C_SOURCE) $(INSTALL_DIR)/src
	cp -r $(MAN_FILES) $(INSTALL_DIR)/man
	cp -r $(R_SOURCE) $(INSTALL_DIR)/R
#	cp install.R $(INSTALL_DIR)
	cp $(INSTALL_DIR)/src/Makefile  $(INSTALL_DIR)/src/Makefile.win
	mkdir $(INSTALL_DIR)/inst
	mkdir $(INSTALL_DIR)/inst/Docs
	if test -n "${DOCS}" ; then cp $(DOCS)  $(INSTALL_DIR)/inst/Docs ; fi
	cp -r examples $(INSTALL_DIR)/inst
	find $(INSTALL_DIR) -name '*~' -exec rm {} \;

PWD=$(shell pwd)

zip: install
	(cd $(RHOME)/library ; zip -r $(ZIP_FILE) $(PKG_NAME); mv $(ZIP_FILE) $(PWD))

binary:  package
	(cd $(RHOME)/src/library ; Rcmd build --binary $(PKG_NAME); mv $(ZIP_FILE) $(PWD))

source: package
	(cd $(RHOME)/src/library ; tar zcf $(TAR_SRC_FILE) $(PKG_NAME); mv $(TAR_SRC_FILE) $(PWD))

#	(cd $(RHOME)/src/library ; zip -r $(ZIP_SRC_FILE) $(PKG_NAME); mv $(ZIP_SRC_FILE) $(PWD))

install: package
	(cd $(RHOME)/src/gnuwin32 ; make pkg-$(PKG_NAME))

check: package
	(cd $(RHOME)/src/library ; Rcmd check $(PKG_NAME))

file:
	@echo "${PKG_FILES}"

Docs/%.html: Docs/%.xml
	$(MAKE) -C Docs $(@F)
