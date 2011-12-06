OMEGA_HOME=../../../..


VERSION_INFO=../../VersionInfo

include $(OMEGA_HOME)/Config/GNUmakefile.SWebPackageConfig
include $(OMEGA_HOME)/R/Config/RPackageWebConfig

$(PKG_TAR_FILE):
	touch $(PKG_TAR_FILE)

PKG_TAR_FILE=

CSS_FILES=$(OMEGA_HOME)/Docs/OmegaTech.css

AUX_DOCS=$(BUGS) $(TODO) $(FAQ) $(CHANGES)

# SCREEN_SHOTS=$(wildcard ScreenShots/*jpg)

DOCS=../../inst/doc/event.html
FAQ=../../FAQ.html
TODO=

install: index.html  installDocs
	scp -r index.html ../../examples  $(OTHER_DOCS) $(CSS_FILES) $(AUX_DOCS) $(TARGET_DIR)

#install: index.html pkg/$(PKG_NAME)/inst/Docs index.html
#	scp -r index.html examples  $(OTHER_DOCS) pkg/$(PKG_NAME)/inst/Docs $(CSS_FILES) $(AUX_DOCS) $(PKG_TAR_FILE) $(TARGET_DIR)


installDocs: $(DOCS)
	scp $(DOCS) $(TARGET_DIR)

#	$(update-r-mirror)

include $(OMEGA_HOME)/Config/GNUmakefile.SWebPackageRules