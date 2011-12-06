#include <windows.h>
#include <oaidl.h>
#include <oleauto.h>

#include <olectl.h>

#ifdef ERROR
#undef ERROR
#endif

extern "C" {
  //#include <Defn.h>
#include <Rinternals.h>
#include <Rdefines.h>
}

//XXX
#define RDCOMEVENTS_VERBOSE 0
#define REFERENCE_ENUM 1


#include "converters.h"
#include "RUtils.h"

extern void COMError(HRESULT hr);
extern HRESULT checkErrorInfo(IUnknown *obj, HRESULT status, SEXP *serr);

extern "C" {
  SEXP R_getEnumConnectionPoints(SEXP s_obj);
  SEXP R_getConnectionInterface(SEXP s_con);
  SEXP R_AdviseConnectionPoint(SEXP s_con, SEXP server);
  SEXP R_UnadviseConnectionPoint(SEXP s_con, SEXP id);
  SEXP R_FindConnectionPoint(SEXP s_obj, SEXP s_iid);
  SEXP RDCOM_ProcessEvents();
}

int R_getCLSIDFromString(SEXP className, CLSID *classId);

SEXP
R_getEnumConnectionPoints(SEXP s_obj)
{
 IUnknown *disp = (IUnknown *) getRDCOMReference(s_obj);
 IConnectionPointContainer *con;

 HRESULT hr;
 SEXP ans, els;

               /* Get the connection point container */
 hr = disp->QueryInterface(IID_IConnectionPointContainer, (void **) &con);

 if(hr != S_OK || !con) {
#ifdef RDCOMEVENTS_VERBOSE
   fprintf(stderr, "Can't query the IConnectionPointContainer interface\n");fflush(stderr);
#endif
   COMError(hr);
 }
 
 con->AddRef();  /* Should we do this? XXX*/

 IEnumConnectionPoints *ppEnum = NULL;
       /* Now get an enumerator from which we can get the individual connection points. */
 hr = con->EnumConnectionPoints(&ppEnum); 

 if(hr != S_OK || !ppEnum) {
#if RDCOMEVENTS_VERBOSE 
    fprintf(stderr, "Error in EnumConnectionPoints method for container\n");
#endif
     con->Release();
     COMError(hr);
 }

 ppEnum->AddRef(); /* And this one too ? XXX */


 hr = S_OK;
 UINT n = 0, i, nprotect = 0;
 long unsigned int which = 0, count = 0;
 unsigned long val = 1;
 unsigned long numRequested = 1;        /* Do 1 at a time for the present. */

 IConnectionPoint *point[100];
 PROTECT(ans = NEW_CHARACTER(100));nprotect++;
 PROTECT(els = NEW_LIST(100));nprotect++;

 /* Added in an effort to avoid crashing or getting errors.*/
 ppEnum->Reset();

 /* Loop over the enumeration, getting blocks at a time. Then for each
    element in the block, add it to the S list. */
 IConnectionPoint *prev = NULL;
 while(hr == S_OK && val > 0) {

#if RDCOMEVENTS_VERBOSE 
   fprintf(stderr, "Getting next set\n");fflush(stderr);
#endif

  hr = ppEnum->Next(numRequested, point, &val);

#if RDCOMEVENTS_VERBOSE 
   fprintf(stderr, "Got %d (requested %d)  %p\n", (int) val, numRequested, (void *)point[0]);fflush(stderr);
#endif

   /* If we get an error and already have some values, return those. 
      If this is not appropriate, throw in a warning. */
   if(((hr != S_OK && val == 0) || (val == 1 && point[0] == NULL)) /* && count */
        || val < numRequested  || (prev && prev == point[0])  ) {
     if(Rf_length(ans) < count) {
       PROTECT(SET_LENGTH(ans, count)); nprotect++;
       PROTECT(SET_LENGTH(els, count)); nprotect++;
     }
     break;
   }

    prev = point[0];

   if(hr != S_OK) {
     SEXP err;
#ifdef RDCOMEVENTS_VERBOSE
     fprintf(stderr, "Error in geting enumeration %d\n", hr);fflush(stderr);
#endif

     if(checkErrorInfo(ppEnum, hr, &err) == S_OK) {
#ifdef RDCOMEVENTS_VERBOSE
       fprintf(stderr, "Got Error Info from object\n");fflush(stderr);
       Rf_PrintValue(err);
#endif
     }
     con->Release();
     COMError(hr);
   }

   IID iid;
   BSTR ostr;
   count += val;
#if RDCOMEVENTS_VERBOSE
   fprintf(stderr, "Count %d, %d %d\n", (int) count, (int) val, Rf_length(ans));fflush(stderr);
#endif
   if(Rf_length(ans) < count) {
     PROTECT(SET_LENGTH(ans, count)); nprotect++;
     PROTECT(SET_LENGTH(els, count)); nprotect++;
   }


   for(i = 0; i < val; i++) {
     SEXP tmp;
     if(point[i]) {
       hr = point[i]->GetConnectionInterface(&iid);
       hr = StringFromCLSID(iid, &ostr);
       SET_STRING_ELT(ans, which, COPY_TO_USER_STRING(FromBstr(ostr)));
       SET_VECTOR_ELT(els, which, tmp = R_createRCOMUnknownObject(point[i], "IConnectionPoint"));
       which++;
       /* Need to release them when we no longer need the results. point[i]->Release(); */
     } else {
       fprintf(stderr, "Skipping connection point %d\n", i);
     }
   }
 }


 if(Rf_length(ans) != which) {
     PROTECT(SET_LENGTH(ans, count)); nprotect++;
     PROTECT(SET_LENGTH(els, count)); nprotect++;
 }

 ppEnum->Release();
 con->Release();

 if(Rf_length(els)) 
   SET_NAMES(els, ans);

 UNPROTECT(nprotect);
 return(els);
}


SEXP
R_FindConnectionPoint(SEXP s_obj, SEXP s_iid)
{
  HRESULT hr;
  IConnectionPoint *cpoint;
  IConnectionPointContainer *con;
  SEXP ans;
  CLSID iid, *ptr;
  IUnknown *disp = (IUnknown *) getRDCOMReference(s_obj);
  ITypeInfo *info = NULL;
  TYPEATTR *attr = NULL;

  hr = disp->QueryInterface(IID_IConnectionPointContainer, (void **) &con);
  if(hr != S_OK) {
#ifdef RDCOMEVENTS_VERBOSE
    fprintf(stderr, "Can't query interface in call to findConnectionPoint\n");
#endif
    COMError(hr);
  }

  if(IS_CHARACTER(s_iid)) {
    hr = R_getCLSIDFromString(s_iid, &iid);
    ptr = &iid;
    if(hr != S_OK) {
#ifdef RDCOMEVENTS_VERBOSE
      fprintf(stderr, "Can't get CLSID in call to findConnectionPoint for %s\n",
	      CHAR(STRING_ELT(s_iid, 0)));
#endif
      con->Release();
      COMError(hr);
    } 
  } else if(TYPEOF(s_iid) == EXTPTRSXP) {
    info = (ITypeInfo *) R_ExternalPtrAddr(s_iid);
    info->GetTypeAttr(&attr);
    ptr = &attr->guid;
  }

  hr = con->FindConnectionPoint(*ptr, &cpoint);

  if(info)
     info->ReleaseTypeAttr(attr);

  if(hr != S_OK) {
#ifdef RDCOMEVENTS_VERBOSE
    fprintf(stderr, "Can't find connection point in call to findConnectionPoint\n");
#endif
    con->Release();
    COMError(hr);
  } 

  ans = R_createRCOMUnknownObject(cpoint, "IConnectionPoint");
  return(ans);
}


/*
 This does not appear to be used.
 What it seems to do is get the UUID of the interface directly from
 the connection point. And we do indeed need that if we want to look
 up information from the type library given the connection point.
 */
SEXP
R_getConnectionInterface(SEXP s_con)
{
 IConnectionPoint *disp = (IConnectionPoint *) getRDCOMReference(s_con);
 IID iid;
 BSTR ostr;
 HRESULT hr;
 SEXP ans;
     hr = disp->GetConnectionInterface(&iid);
     hr = StringFromCLSID(iid, &ostr);
     PROTECT(ans = NEW_CHARACTER(1));
     SET_STRING_ELT(ans, 0, COPY_TO_USER_STRING(FromBstr(ostr)));
     UNPROTECT(1);
 return(ans);
}

IUnknown *
R_getServerInstance(SEXP server)
{
  IUnknown *obj; 
  SEXP ref =  GET_SLOT(server, Rf_install("ref"));
  obj = (IUnknown *) R_ExternalPtrAddr(ref);
  if(!obj) {
   PROBLEM "NULL reference passed as DCOM server"
   ERROR;
  }
  return(obj);
}

SEXP
R_AdviseConnectionPoint(SEXP s_con, SEXP server)
{
 IUnknown *unk;
 unk = (IUnknown *) R_getServerInstance(server);
 /* Should we use unk->QueryInterface(&IID_Unknown)*/

 IConnectionPoint *disp = (IConnectionPoint *) getRDCOMReference(s_con);
 HRESULT hr;
 DWORD cookie;
 hr = disp->Advise(unk, &cookie);
 if(hr != S_OK) {
   COMError(hr);
 } else {
   errorLog("Connected to connection point via Advise %p %d  (%X)\n", unk, (int) cookie, hr);
 }

 SEXP ans = NEW_INTEGER(1); 
 INTEGER_DATA(ans)[0] = cookie;
 return(ans);
}

SEXP
R_UnadviseConnectionPoint(SEXP s_con, SEXP id)
{
 IConnectionPoint *disp = (IConnectionPoint *) getRDCOMReference(s_con);
 HRESULT hr;
 DWORD cookie;

 cookie = (DWORD) INTEGER_DATA(id)[0];
 hr = disp->Unadvise(cookie);
 if(hr != S_OK) {
   COMError(hr);
 }

 return(R_NilValue);
}

extern "C" {
   void R_ProcessEvents(void);
}

SEXP
RDCOM_ProcessEvents()
{
#if 0
  unsigned long ctr = 0;
  while(PeekMessage(&msg, 0, 0, 0, PM_NOREMOVE) {
      ctr++;
      ProcessEvent();
  }

  return ScalarReal(ctr);
#endif
   R_ProcessEvents();
   return(R_NilValue);
}

