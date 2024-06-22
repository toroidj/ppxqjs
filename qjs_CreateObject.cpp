/*-----------------------------------------------------------------------------
	Paper Plane xUI QuickJS Script Module							(c)TORO
	PPx.CreateObject 関係

todo
	EnumConnectionPoints の2以上に対応
	safearray 対応
	ref 対応
-----------------------------------------------------------------------------*/
#define STRICT
#define UNICODE
#include <stddef.h>
#include <windows.h>
#include <dispex.h>
#include "quickjs.h"
#include "TOROWIN.H"
#include "PPCOMMON.H"
#include "ppxqjs.h"

JSClassID ClassID_Dispatch = 0;
JSClassID ClassID_EnumVARIANT = 0;

#ifndef RELESE
int CountDisp = 0, CountDispRelease = 0, CountEvent = 0, CountEventRelease = 0;
#endif

#define DISPNAME_MAXLEN 96
JSValue JS_NewDispatch(JSContext *ctx, IDispatch *pDisp, const char *prefix);
JSValue JS_NewObjectFromVARIANT(JSContext *ctx, VARIANT *var);

double fixtimezone = 0.;
BOOL gottimezone = FALSE;;

typedef union {
	int64_t datetime;
	FILETIME ftime;
} FILETIME64;

void GetTimezone(void) // ゾーン切り替えに対応しない
{
	FILETIME64 gmt , utc;

	gmt.datetime = 864000000000.; // 1日
	LocalFileTimeToFileTime(&gmt.ftime, &utc.ftime);
	fixtimezone = (double)utc.datetime / 10000.0  - 86400000.;
	gottimezone = TRUE;
}

BOOL GetRegStringz(HKEY hKey, const WCHAR *path, const WCHAR *name, WCHAR *dest, size_t destlen)
{
	HKEY HK;
	DWORD typ, siz;
//	TCHAR buf[VFPS];

	if ( RegOpenKeyEx(hKey, path, 0, KEY_READ, &HK) == ERROR_SUCCESS ){
		siz = destlen * sizeof(TCHAR);
		if ( RegQueryValueEx(HK, name, NULL, &typ, (LPBYTE)dest, &siz) == ERROR_SUCCESS ){
			/*
			if ( (typ == REG_EXPAND_SZ) && (siz < sizeof(buf)) ){

				tstrcpy(buf, dest);
				ExpandEnvironmentStrings(buf, dest, destlen);
			}
			*/
			RegCloseKey(HK);
			return TRUE;
		}
		RegCloseKey(HK);
	}
	return FALSE;
}


const char *VTnamelist[] = {
	"empty", "null", "short", "long", "float", "double", "current", "date",
	"string", "object", "error", "bool", "variant", "dataobj", "desimal", "unknown15",
	"int_8", "uint_8", "ushort", "uint", "int_64", "uint_64", "INT", "UINT",
	"void", "HRESULT", "PTR", "array", "carray", "user", "LPSTR", "LPWSTR"
	};

const char *GetVTname(VARTYPE vt, char *buf)
{
	if ( vt & 0x2000 ) return "array";
	if ( (vt >= 0) && (vt < _countof(VTnamelist)) ) return VTnamelist[vt];
	sprintf(buf, "(unknownID:%d)", vt);
	return buf;
}

typedef struct {
	IEnumVARIANT* pEnum;
	BOOL done;
} ENUMVARIANTINFO;
#define GetEnumVARIANTInfo(this_obj) ((ENUMVARIANTINFO *)JS_GetOpaque(this_obj, ClassID_EnumVARIANT))

JSValue GetEnumVARIANTDone(JSContext *ctx, JSValueConst this_obj)
{
	ENUMVARIANTINFO *info = GetEnumVARIANTInfo(this_obj);
	return JS_NewBool(ctx, info->done);
}

JSValue EnumVARIANTNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENUMVARIANTINFO *info = GetEnumVARIANTInfo(this_obj);
	DWORD c;
	VARIANT var;

	if ( info->pEnum == NULL ) return JS_EXCEPTION;
	VariantInit(&var);
	if ( SUCCEEDED(info->pEnum->Next(1, &var, &c)) ){
		info->done = FALSE;
		JSValue jsvar = JS_NewObjectFromVARIANT(ctx, &var);
		if ( var.vt != VT_EMPTY ){
			VariantClear(&var);
			JS_SetPropertyStr(ctx, this_obj, "value", jsvar);
			return JS_DupValue(ctx, this_obj);
		}
	}
	info->done = TRUE;
	JS_SetPropertyStr(ctx, this_obj, "value", JS_UNDEFINED);
	return JS_DupValue(ctx, this_obj);
}

JSValue EnumVARIANTFree(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENUMVARIANTINFO *info = GetEnumVARIANTInfo(this_obj);
	if ( info->pEnum != NULL ){
		info->pEnum->Release();
		info->pEnum = NULL;
	}
	return JS_UNDEFINED;
}

static const JSCFunctionListEntry EnumVARIANTFunctionList[] = {
	JS_CFUNC_DEF("next", 0, EnumVARIANTNext),
//	JS_CGETSET_DEF("value", GetEnumVARIANTValue, NULL),
	JS_CGETSET_DEF("done", GetEnumVARIANTDone, NULL),
	JS_CFUNC_DEF("free", 0, EnumVARIANTFree),
};

void EnumVARIANTFinalizer(JSRuntime *rt, JSValue this_obj)
{
	ENUMVARIANTINFO *info = GetEnumVARIANTInfo(this_obj);
	if ( info->pEnum != NULL ) info->pEnum->Release();
	free(info);
}

JSClassDef EnumVARIANTClassDef = {
	"_DispEnuM_", // class_name
	EnumVARIANTFinalizer, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};

struct DISPATCHSTRUCT;

class CEventSink : public IDispatch
{
private:
	DISPATCHSTRUCT *m_disps;
	IConnectionPoint *m_ConnectPoint;
	ITypeInfo *m_EventType;
	int m_refCount;
	char m_prefix[DISPNAME_MAXLEN];
	DWORD m_ConnectID;
	int m_LogFire;

public:
	CEventSink(DISPATCHSTRUCT *disps, const char *prefix);
	virtual ~CEventSink();
//----------------------------------------------------- IUnknown
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj);
//-------------------------------------------------- IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *iTInfo);
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames, LCID, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lc, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO * ei, UINT *puArgErr);
//-------------------------------------------------- init
	void Connect(IConnectionPoint *pPoint);
	void Disconnect();
	void SetEventInfo(ITypeInfo *pEventType);
	void SetEventTrace(int trace);
};

struct DISPATCHSTRUCT {
	FREECHAIN chain;
	JSContext *ctx;
	IDispatch *pDisp;
	CEventSink *pEventSink;
};

#define GetDispatchPtr(thisobj) (reinterpret_cast<DISPATCHSTRUCT *>(JS_GetOpaque(thisobj, ClassID_Dispatch)))
#define GetDispatchDisp(thisobj, pIDisp) {\
	DISPATCHSTRUCT *disps = GetDispatchPtr(thisobj);\
	if ( disps == NULL ) return JS_EXCEPTION;\
	(pIDisp) = disps->pDisp;\
	if ( (pIDisp) == NULL ) return JS_EXCEPTION;\
}

JSValue DispatchIterator(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	DISPPARAMS params;
	HRESULT result;
	VARIANT resultvalue;
	IDispatch *pDisp;

	GetDispatchDisp(this_obj, pDisp);

	params.rgvarg = NULL;
	params.rgdispidNamedArgs = NULL;
	params.cArgs = 0;
	params.cNamedArgs = 0;

	result = pDisp->Invoke(DISPID_NEWENUM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &resultvalue, NULL, NULL);
	if ( SUCCEEDED(result) ){
		JSRuntime *rt = JS_GetRuntime(ctx);

		if ( ClassID_EnumVARIANT == 0 ) JS_NewClassID(&ClassID_EnumVARIANT);
		if ( !JS_IsRegisteredClass(rt, ClassID_EnumVARIANT) ){ // クラスを登録
			JS_NewClass(rt, ClassID_EnumVARIANT, &EnumVARIANTClassDef);
			JSValue ProtoClass = JS_NewObject(ctx);
			JS_SetPropertyFunctionList(ctx, ProtoClass, EnumVARIANTFunctionList, _countof(EnumVARIANTFunctionList));
			JS_SetClassProto(ctx, ClassID_EnumVARIANT, ProtoClass);
		}
		JSValue newobj = JS_NewObjectClass(ctx, ClassID_EnumVARIANT);
		ENUMVARIANTINFO *newinfo = static_cast<ENUMVARIANTINFO *>(malloc(sizeof(ENUMVARIANTINFO)));
		newinfo->done = FALSE;
		newinfo->pEnum = reinterpret_cast<IEnumVARIANT *>(resultvalue.punkVal);
		JS_SetOpaque(newobj, newinfo);
		return newobj;
	}
	JS_ThrowReferenceError(ctx, "invoke enum not support");
	return JS_EXCEPTION;
}

void reportf(JSContext *ctx, const char *message, ...)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	char buf[0x400];
	WCHAR bufW[0x400];
	t_va_list argptr;

	t_va_start(argptr, message);
	wvsprintfA(buf, message, argptr);
	t_va_end(argptr);
	buf[sizeof(buf) / sizeof(char) - 1] = '\0';
	MultiByteToWideChar(CP_UTF8, 0, buf, -1, bufW, 0x400);
	bufW[sizeof(buf) / sizeof(WCHAR) - 1] = '\0';
	ppxa->Function(ppxa, PPXCMDID_REPORTTEXT, bufW);
}

void REPORT(JSContext *ctx,ThSTRUCT *TH)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	ppxa->Function(ppxa, PPXCMDID_REPORTTEXT, (WCHAR *)TH->bottom);
	THFree(TH);
}


JSValue JS_NewObjectFromVARIANT(JSContext *ctx, VARIANT *var)
{
	JSValue jsvar;

	switch ( var->vt ){
		case VT_EMPTY:
			jsvar = JS_UNDEFINED;
			break;

		case VT_NULL:
			jsvar = JS_NULL;
			break;

		case VT_I2:
			jsvar = JS_NewInt32(ctx, static_cast<int>(var->iVal));
			break;

		case VT_I4:
			jsvar = JS_NewInt32(ctx, var->intVal);
			break;

		case VT_R4:
			jsvar = JS_NewFloat64(ctx, static_cast<double>(var->fltVal));
			break;

		case VT_R8:
			jsvar = JS_NewFloat64(ctx, var->dblVal);
			break;

//		case VT_CY:

		case VT_BSTR:
			jsvar = JS_NewStringW(ctx, var->bstrVal);
			break;

		case VT_DISPATCH:
			var->pdispVal->AddRef();
			#ifndef RELESE
				CountDisp++;
			#endif
			jsvar = JS_NewDispatch(ctx, var->pdispVal, NULL);
			break;

		case VT_ERROR:
			jsvar = JS_EXCEPTION;
			break;

		case VT_BOOL:
			jsvar = JS_NewBool(ctx, var->boolVal);
			break;

//		case VT_VARIANT:
//		case VT_UNKNOWN:
//		case VT_UI1:

		case VT_DATE:
			if ( !gottimezone ) GetTimezone();
			// 1900/01/01-2 → 1970/01/01 に変換
			jsvar = JS_NewInt64(ctx, static_cast<__int64>((var->date - 25569.) * 86400000 + fixtimezone));
			break;

		default: // その他は文字列扱い
			#ifndef RELESE
				reportf(ctx, "\r\nunknown VARIANT %d\r\n", var->vt);
			#endif
			if ( SUCCEEDED(::VariantChangeType(var, var, 0, VT_BSTR)) ){
				jsvar = JS_NewStringW(ctx, var->bstrVal);
			}else{
				JS_ThrowReferenceError(ctx, "unknown VARIANT(%d)", var->vt);
				jsvar = JS_NULL;
			}
	}
	return jsvar;
}

#define MAXVALBUF 5
static JSValue DispatchInvoke(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv, unsigned short wFlags, DISPID id)
{
	DISPPARAMS dispparam;
	VARIANT inputbuf[MAXVALBUF], *input = inputbuf, output;
	WCHAR bufW[CMDLINESIZE];
	HRESULT result;
	IDispatch *pDisp;

	GetDispatchDisp(this_obj, pDisp);

	if ( argc < 0 ) return JS_EXCEPTION;
	if ( argc > MAXVALBUF ) input = new VARIANT[argc];

	for ( int c = argc - 1; c >= 0; c--){
		#define disp_c (argc - 1 - c)
		switch ( JS_VALUE_GET_TAG(argv[c]) ){
			case JS_TAG_INT:
				input[disp_c].vt = VT_I4;
				JS_ToInt32(ctx, &input[disp_c].intVal, argv[c]);
				break;

			case JS_TAG_BOOL:
				input[disp_c].vt = VT_BOOL;
				input[disp_c].boolVal = JS_ToBool(ctx, argv[c]);
				break;

			case JS_TAG_NULL:
				input[disp_c].vt = VT_NULL;
				break;

			case JS_TAG_UNDEFINED:
			case JS_TAG_UNINITIALIZED:
				input[disp_c].vt = VT_EMPTY;
				break;

//			case JS_TAG_CATCH_OFFSET:

			case JS_TAG_EXCEPTION:
//				input[disp_c].vt = VT_ERROR;
				input[disp_c].vt = VT_NULL;
				break;

			case JS_TAG_FLOAT64:
				input[disp_c].vt = VT_R8;
				JS_ToFloat64(ctx, &input[disp_c].dblVal, argv[c]);
				break;

			default: // その他は文字列扱い
			#ifndef RELESE
				reportf(ctx, "unknown JSValue(Arg:%d, Tag:%d)", c, JS_VALUE_GET_TAG(argv[c]));
			#endif

			case JS_TAG_OBJECT:
			case JS_TAG_STRING:
				WCHAR *param;

				param = GetJsLongString(ctx, argv[c], bufW);
				input[disp_c].vt = VT_BSTR;
				input[disp_c].bstrVal = ::SysAllocString(param);
				FreeJsLongString(param, bufW);
				break;
		}
	}

	VariantInit(&output);

	dispparam.rgvarg = input;
	dispparam.cArgs = argc;

	DISPID dispidName = DISPID_PROPERTYPUT;
	if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF) ) {
		dispparam.cNamedArgs = 1;
		dispparam.rgdispidNamedArgs = &dispidName;
	}else{
		dispparam.cNamedArgs = 0;
		dispparam.rgdispidNamedArgs = NULL;
	}

	result = pDisp->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT, wFlags, &dispparam, &output, NULL, NULL);

	for ( int c = 0 ; c < argc; c++ ) VariantClear(&input[c]);
	if ( input != inputbuf ) delete[] input;

	if ( FAILED(result) ){
		const char *msg;

		switch(result){
			case DISP_E_BADPARAMCOUNT:
				msg = "invoke: number of parameter(%d) is different";
				result = (HRESULT)(argc - 1);
				break;

			case DISP_E_BADVARTYPE:
				msg = "invoke: valid variant type";
				break;

			case DISP_E_EXCEPTION:
				msg = "invoke: exception";
				break;

			case DISP_E_MEMBERNOTFOUND:
				msg = "invoke: call type error";
				break;

			case DISP_E_OVERFLOW:
				msg = "invoke: One of the parameter could not be coerced to the specified type";
				break;

			case DISP_E_PARAMNOTFOUND:
				msg = "invoke: One of the parameter does not correspond";
				break;

			case DISP_E_TYPEMISMATCH:
				msg = "invoke: could not be coerced.";
				break;

			case DISP_E_PARAMNOTOPTIONAL:
				msg = "invoke: A required parameter was omitted.";
				break;

			default:
				msg = "invoke error(%x)";
				break;
		}
		if ( msg != NULL ){
			JS_ThrowReferenceError(ctx, msg, (int)result);
			return JS_EXCEPTION;
		}
	}

	JSValue jsvar = JS_NewObjectFromVARIANT(ctx, &output);
	VariantClear(&output);
	return jsvar;
}

JSValue DispatchFunction(JSContext *ctx, JSValueConst this_val, int argc, JSValueConst *argv, int magic)
{
	return DispatchInvoke(ctx, this_val, argc, argv, DISPATCH_METHOD, magic);
}

JSValue DispatchGetter(JSContext *ctx, JSValueConst this_obj, int magic)
{
	return DispatchInvoke(ctx, this_obj, 0, NULL, DISPATCH_PROPERTYGET, magic);
}

JSValue DispatchSetter(JSContext *ctx, JSValueConst this_obj, JSValueConst val, int magic)
{
	return DispatchInvoke(ctx, this_obj, 1, &val, DISPATCH_PROPERTYPUT, magic);
}

DISPID GetDispID(JSContext *ctx, IDispatch *pDisp, JSValueConst &name)
{
	DISPID id;

	if ( JS_IsNumber(name) ){
		id = DISPID_UNKNOWN;
		JS_ToInt32(ctx, (int *)&id, name);
	}else{
		WCHAR bufW[CMDLINESIZE], *nameptr = bufW;

		GetJsShortString(ctx, name, bufW);
		if ( FAILED(pDisp->GetIDsOfNames(IID_NULL, &nameptr, 1, LOCALE_SYSTEM_DEFAULT, &id)) ){
			id = DISPID_UNKNOWN;
		}
	}
	return id;
}

static JSValue DispatchInvokeRaw(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv, int magic, IDispatch *pDisp)
{
	DISPID id;

	if ( argc < 1 ){
		JS_ThrowRangeError(ctx, "__Invoke__(name, param1, param2...)");
		return JS_EXCEPTION;
	}
	id = GetDispID(ctx, pDisp, argv[0]);
	if ( id == DISPID_UNKNOWN ){
		JS_ThrowReferenceError(ctx, "Invoke name not found");
		return JS_EXCEPTION;
	}
	return DispatchInvoke(ctx, this_obj, argc - 1, argv + 1, static_cast<unsigned short>(magic), id);
}

void DispatchRelease(DISPATCHSTRUCT *disps)
{
	if ( disps == NULL ) return;
#ifndef RELESE
	CountDispRelease++;
#endif
	if ( disps->pEventSink != NULL ){
		disps->pEventSink->Disconnect();
		disps->pEventSink->Release();
	}
	disps->pDisp->Release();
	free(disps);
}

JSValue DispatchFree(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	DISPATCHSTRUCT *disps = GetDispatchPtr(this_obj);
	if ( disps != NULL ){
		DispatchRelease(disps);
		JS_SetOpaque(this_obj, NULL);
	}
	return JS_UNDEFINED;
}

static JSValue DispatchToString(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	VARIANT var;
	JSValue jsvar;

	GetDispatchDisp(this_obj, var.pdispVal);
	var.vt = VT_DISPATCH;

	var.pdispVal->AddRef();
	if ( SUCCEEDED(::VariantChangeType(&var, &var, 0, VT_BSTR)) ){
		jsvar = JS_NewStringW(ctx, var.bstrVal);
	}else{
		jsvar = JS_NULL;
	}
	VariantClear(&var);

	return jsvar;
}


void LogPrintf(ThSTRUCT *TH, const char *message, ...)
{
	char buf[0x400];
	WCHAR bufW[0x400];
	t_va_list argptr;

	t_va_start(argptr, message);
	wvsprintfA(buf, message, argptr);
	t_va_end(argptr);
	buf[sizeof(buf) / sizeof(char) - 1] = '\0';
	MultiByteToWideChar(CP_UTF8, 0, buf, -1, bufW, 0x400);
	bufW[sizeof(buf) / sizeof(WCHAR) - 1] = '\0';
	THCatStringW(TH, bufW);
}

void LogPrintWchar(ThSTRUCT *TH, const char *format, WCHAR *str)
{
	if ( str == NULL ){
		THCatStringW(TH, L"(null)");
	}else{
		char buf[200];
		if ( 0 >= WideCharToMultiByte(CP_UTF8, 0, str, -1, buf, 200, NULL, NULL) ){
			buf[100] = '\0';
		}
		LogPrintf(TH, format, buf);
	}
}

void LogPrintParamName(ThSTRUCT *TH, const char *format, BSTR *list, unsigned int getnames, unsigned int index)
{
	BSTR str;

	str = NULL;
	if ( index < getnames ){
		if ( list != NULL ) str = list[index];
	}

	if ( str == NULL ){
	//	reportf(ctx, format, "(null)");
	}else{
		char buf[200];
		if ( 0 >= WideCharToMultiByte(CP_UTF8, 0, str, -1, buf, 200, NULL, NULL) ){
			buf[100] = '\0';
		}
		LogPrintf(TH, format, buf);
	}
}

void PrintFuncDesc(ThSTRUCT *TH, ITypeInfo *iType, FUNCDESC *fdesc)
{
	BSTR *names;
	unsigned int getnames;

	char vtbuf[1000];

	THCatStringW(TH, L"  ");

	getnames = fdesc->cParams + 1;
	if ( fdesc->invkind & (DISPATCH_METHOD | DISPATCH_PROPERTYGET) ) getnames++;
	names = new BSTR[getnames];

	if ( FAILED(iType->GetNames(fdesc->memid, names, getnames, &getnames)) ){
		getnames = 0;
	}

	if ( fdesc->invkind & (DISPATCH_METHOD | DISPATCH_PROPERTYGET) ){
		if ( (fdesc->invkind & DISPATCH_METHOD) && (fdesc->elemdescFunc.tdesc.vt == VT_VOID) ){
			THCatStringW(TH, L"void ");
		}else{
			LogPrintf(TH, "%s ", GetVTname(fdesc->elemdescFunc.tdesc.vt, vtbuf));
			LogPrintParamName(TH, "%s ", names, getnames, fdesc->cParams + 1);
			THCatStringW(TH, L"= ");
		}
	}

	LogPrintParamName(TH, "%s ", names, getnames, 0);

	if ( fdesc->invkind & DISPATCH_METHOD ){
		THCatStringW(TH, L"(");
	}
	if ( fdesc->invkind & DISPATCH_PROPERTYPUT ){
		THCatStringW(TH, L" =");
	}
	if ( fdesc->cParams > 0 ){
		for ( int paramc = 0 ; paramc < fdesc->cParams ; paramc++ ){
			TYPEDESC t;
			if ( paramc >= (fdesc->cParams - fdesc->cParamsOpt) ){
				THCatStringW(TH, L" [option]");
			}
			if ( fdesc->invkind & DISPATCH_METHOD ){
				LogPrintf (TH, " %d:",paramc + 1);
			}else{
				THCatStringW(TH, L" ");
			}
			t = fdesc->lprgelemdescParam[paramc].tdesc;
			while ( t.vt == VT_SAFEARRAY || t.vt == VT_PTR ) {
				t = *t.lptdesc;
			}

			LogPrintf(TH, "%s", GetVTname(t.vt, vtbuf));
			LogPrintParamName(TH, " %s", names, getnames, paramc + 1);
			if ( (paramc + 1) < fdesc->cParams ) THCatStringW(TH, L",");
		}
	}
	if ( fdesc->invkind & DISPATCH_METHOD ){
		THCatStringW(TH, L")");
	}
	LogPrintf(TH, " [%d]\r\n", fdesc->memid);

	for ( unsigned int i = 0; i < getnames; i++ ){
		::SysFreeString(names[i]);
	}
	delete[] names;

	BSTR docstr;
	if ( SUCCEEDED(iType->GetDocumentation(fdesc->memid , NULL, &docstr, NULL, NULL)) ){
		if ( docstr != NULL ) LogPrintWchar(TH, "  -- %s\r\n", docstr);
		::SysFreeString(docstr);
	}

	THCatStringW(TH, L"\r\n");
}

// 列挙用 ITypeInfo とその TYPEATTR を取得
BOOL GetTypeInfoAttr(IDispatch *pDisp, ITypeInfo **pType, TYPEATTR **attr, ITypeInfo **pEventType, IID *eventID)
{
	ITypeLib *pTLib;
	UINT libindex;

	if ( FAILED(pDisp->GetTypeInfo(0, LOCALE_USER_DEFAULT, pType)) ){
		return FALSE;
	}
	if ( FAILED((*pType)->GetTypeAttr(attr)) ){
		(*pType)->Release();
		return FALSE;
	}

	// pType には承継元の情報が含まれていない可能性があるため、ライブラリも確認
	if ( SUCCEEDED((*pType)->GetContainingTypeLib(&pTLib, &libindex)) ){
	//	int libmax = pTLib->GetTypeInfoCount();
		ITypeInfo *pLibType;

		if ( eventID != NULL ){
			if ( FAILED(pTLib->GetTypeInfoOfGuid(*eventID, pEventType))) {
				*pEventType = NULL;
			}
		}

		if ( SUCCEEDED(pTLib->GetTypeInfo(libindex, &pLibType)) ){
			TYPEATTR *libattr;

			if ( SUCCEEDED(pLibType->GetTypeAttr(&libattr)) ){
				if ( libattr->typekind == TKIND_DISPATCH ){
					if ( ((*attr)->guid == libattr->guid) &&
						 ((*attr)->cFuncs < libattr->cFuncs) ){
						(*pType)->ReleaseTypeAttr(*attr);
						(*pType)->Release();
						*pType = pLibType;
						*attr = libattr;
						pTLib->Release();
						return TRUE;
					}
				}
				pLibType->ReleaseTypeAttr(libattr);
			}
			pLibType->Release();
		}
		pTLib->Release();
	}
	return TRUE;
}

IConnectionPoint *GetEventID(IDispatch *pDisp, IID &eventID)
{
	IConnectionPoint *pPoint = NULL;
	IConnectionPointContainer *pContainer;
	IProvideClassInfo2 *pProvider;
	IEnumConnectionPoints *pEnumPoints;

	if ( FAILED(pDisp->QueryInterface(IID_IConnectionPointContainer,
			reinterpret_cast<void**>(&pContainer))) ){
		return NULL;
	}

	if ( SUCCEEDED(pDisp->QueryInterface(IID_IProvideClassInfo2,
			reinterpret_cast<void**>(&pProvider))) ){
		if ( SUCCEEDED(pProvider->GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID, &eventID)) ){
			if ( FAILED(pContainer->FindConnectionPoint(eventID, &pPoint))){
				pPoint = NULL;
			}
		}
		pProvider->Release();
	}else if ( SUCCEEDED(pContainer->EnumConnectionPoints(&pEnumPoints)) ){
		if ( SUCCEEDED(pEnumPoints->Next(1, &pPoint, 0)) ){
			if( FAILED(pPoint->GetConnectionInterface(&eventID)) ){
				pPoint->Release();
				pPoint = NULL;
			}
		}
		pEnumPoints->Release();
	}
	pContainer->Release();
	return pPoint;
}

void PrintInterfaceName(ThSTRUCT *TH, TYPEATTR *attr)
{
	WCHAR path[100];
	WCHAR name[100];
	StringFromGUID2(attr->guid, name , 100);
	wsprintfW(path, L"Interface\\%s", name);
	GetRegStringz(HKEY_CLASSES_ROOT, path, L"", name, 100);
	LogPrintWchar(TH, " <%s>\r\n", name);
}

JSValue DispatchInfo(JSContext *ctx, IDispatch *pDisp, int argc, JSValueConst *argv, BOOL eventmode)
{
	DISPID id;
	BOOL all;
	ThSTRUCT TH = ThSTRUCT_InitData;

	if ( argc >= 1 ){
		const char *name = JS_ToCString(ctx, argv[0]);

		id = GetDispID(ctx, pDisp, argv[0]);
		if ( id == DISPID_UNKNOWN ){
			JS_FreeCString(ctx, name);
			JS_ThrowReferenceError(ctx, "name not found");
			return JS_EXCEPTION;
		}
		all = FALSE;
		LogPrintf(&TH, "\r\n%s information(dispid: %d)\r\n", name, (int)id);
		JS_FreeCString(ctx, name);
	}else{
		THCatStringW(&TH, L"\r\nall items\r\n");
		all = TRUE;
	}

	ITypeInfo *pType;
	ITypeInfo *pOtherType;
	TYPEATTR *attr;

	if ( eventmode ){
		IID eventID;
		IConnectionPoint *pPoint = GetEventID(pDisp, eventID);
		if ( pPoint != NULL ){
			pPoint->Release();
			if ( GetTypeInfoAttr(pDisp, &pOtherType, &attr, &pType, &eventID) ){
				pOtherType->ReleaseTypeAttr(attr);
				pOtherType->Release();
				pType->GetTypeAttr(&attr);
			}else{

				REPORT(ctx, &TH);
				return JS_UNDEFINED;
			}
		}
	}else{
		if ( GetTypeInfoAttr(pDisp, &pType, &attr, NULL, NULL) == FALSE ){
			REPORT(ctx, &TH);
			return JS_UNDEFINED;
		}
	}

	PrintInterfaceName(&TH, attr);

	unsigned int funcmax = static_cast<unsigned int>(attr->cFuncs);
	for (unsigned int funcno = 0; funcno < funcmax; funcno++ ){
		FUNCDESC *fdesc;

		if ( FAILED(pType->GetFuncDesc(funcno, &fdesc)) ) continue;

		if ( (all || (fdesc->memid == id)) && !(fdesc->wFuncFlags & (FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN)) ){
			PrintFuncDesc(&TH, pType, fdesc);
		}
		pType->ReleaseFuncDesc(fdesc);
	}
	pType->ReleaseTypeAttr(attr);
	pType->Release();
	REPORT(ctx, &TH);
	return JS_UNDEFINED;
}

void InitEventSink(DISPATCHSTRUCT *disps, const char *prefix, IID &eventID)
{
	IConnectionPoint *pPoint = GetEventID(disps->pDisp, eventID);
	if ( pPoint != NULL ){
		#ifndef RELESE
			CountEvent++;
		#endif
		disps->pEventSink = new CEventSink(disps, prefix);
		disps->pEventSink->Connect(pPoint); // pPoint の生存管理は、CEventSink で行う

		InstanceValueStruct *info = GetCtxInfo(disps->ctx);
		Chain_FreeChainObject(info->EventChain, disps->chain, reinterpret_cast<void *>(disps->pEventSink));
	}
}


JSValue DispatchExtra(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	DISPATCHSTRUCT *disps = GetDispatchPtr(this_obj);
	IDispatch *pDisp = disps->pDisp;
	JSValue result = JS_UNDEFINED;
	const char *cmd = JS_ToCString(ctx, argv[0]);

	if ( strcmp(cmd, "get") == 0 ){
		result = DispatchInvokeRaw(ctx, this_obj, argc - 1, argv + 1, DISPATCH_PROPERTYGET, pDisp);
	}else if ( (strcmp(cmd, "set") == 0) || (strcmp(cmd, "put") == 0) ){
		result = DispatchInvokeRaw(ctx, this_obj, argc - 1, argv + 1, DISPATCH_PROPERTYPUT, pDisp);
	}else if ( (strcmp(cmd, "function") == 0) || (strcmp(cmd, "method") == 0) ){
		result = DispatchInvokeRaw(ctx, this_obj, argc - 1, argv + 1, DISPATCH_METHOD, pDisp);
	}else if ( strcmp(cmd, "info") == 0 ){
		result = DispatchInfo(ctx, pDisp, argc - 1, argv + 1, FALSE);
	}else if ( strcmp(cmd, "event") == 0 ){
		result = DispatchInfo(ctx, pDisp, argc - 1, argv + 1, TRUE);
	}else if ( strcmp(cmd, "trace") == 0 ){
		if ( disps->pEventSink != NULL ){
			disps->pEventSink->SetEventTrace( (argc < 2) || JS_ToBool(ctx, argv[1]) );
		}
	}else if ( strcmp(cmd, "connect") == 0 ){
		result = JS_FALSE;
		if ( disps->pEventSink == NULL ){
			const char *prefix = JS_ToCString(ctx, argv[1]);
			IID eventID;

			InitEventSink(disps, (prefix != NULL) ? prefix : "", eventID);
			JS_FreeCString(ctx, prefix);

			if ( prefix != NULL ){
				ITypeInfo *pType;
				ITypeInfo *pEventType = NULL;
				TYPEATTR *attr;

				if ( GetTypeInfoAttr(disps->pDisp, &pType, &attr, &pEventType, &eventID) ){
					if ( pEventType != NULL ){
						disps->pEventSink->SetEventInfo(pEventType);
						result = JS_TRUE;
					}
					pType->ReleaseTypeAttr(attr);
					pType->Release();
				}
			}
		}
	}else if ( strcmp(cmd, "disconnect") == 0 ){
		result = JS_FALSE;
		if ( disps->pEventSink != NULL ){
			if ( disps->pEventSink != NULL ) {
				disps->pEventSink->Disconnect();
				result = JS_TRUE;
			}
		}
	}

	JS_FreeCString(ctx, cmd);
	return result;
}

static const JSCFunctionListEntry DispFunctionList[] = {
	JS_CFUNC_DEF("_free_", 0, DispatchFree),
	JS_CFUNC_DEF("_", 0, DispatchExtra),
	JS_CFUNC_DEF("toString", 0, DispatchToString),
	JS_CFUNC_DEF("Symbol.iterator", 0, DispatchIterator),
	JS_CFUNC_DEF("[Symbol.iterator]", 0, DispatchIterator),
};

void DispatchFinalizer(JSRuntime *rt, JSValue this_obj)
{
	DISPATCHSTRUCT *disps = GetDispatchPtr(this_obj);
	DispatchRelease(disps);
}

JSClassDef DispClassDef = {
	"_IDisp_", // class_name
	DispatchFinalizer, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};
#define FUNCNAMEOFF 4

CEventSink::CEventSink(DISPATCHSTRUCT *disps, const char *prefix)
{
	m_refCount = 1; // addref 1回分
	m_disps = disps;
	m_ConnectPoint = NULL;
	m_EventType = NULL;
	strcpy(m_prefix, prefix);
	m_LogFire = FALSE;
}

CEventSink::~CEventSink()
{
}
//----------------------------------------------------- IUnknown
STDMETHODIMP_(ULONG) CEventSink::AddRef()
{
	return ++m_refCount;
}

STDMETHODIMP_(ULONG) CEventSink::Release()
{
	if (--m_refCount == 0){
		if ( m_ConnectPoint != NULL ) Disconnect();
	#ifndef RELESE
		CountEventRelease++;
	#endif
		delete this;
		return 0;
	}
	return m_refCount;
}

STDMETHODIMP CEventSink::QueryInterface(REFIID riid, void **ppvObj)
{
	if ( (riid == IID_IUnknown) || (riid == IID_IDispatch) ){
		*ppvObj = this;
		AddRef();
		return S_OK;
	}
	*ppvObj = NULL;
	return E_NOINTERFACE;
}

//-------------------------------------------------- IDispatch
STDMETHODIMP CEventSink::GetTypeInfoCount(UINT *iTInfo)
{
	*iTInfo = 0;
	return S_OK;
}

STDMETHODIMP CEventSink::GetTypeInfo(UINT, LCID, ITypeInfo **ppTInfo)
{
	*ppTInfo = NULL;
	return DISP_E_BADINDEX;
}

STDMETHODIMP CEventSink::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames, LCID, DISPID *rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CEventSink::Invoke(DISPID dispIdMember, REFIID riid, LCID lc, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO * ei, UINT *puArgErr)
{
	char funcname[DISPNAME_MAXLEN * 2], eventname[DISPNAME_MAXLEN];

	eventname[0] = '\0';
	if ( m_EventType != NULL ){
		BSTR name;
		UINT c;

		if ( SUCCEEDED(m_EventType->GetNames(dispIdMember, &name, 1, &c)) ){
			if ( 0 >= WideCharToMultiByte(CP_UTF8, 0, name, -1, eventname, 250, NULL, NULL) ){
				eventname[DISPNAME_MAXLEN - 1] = '\0';
			}
			::SysFreeString(name);
		}
	}else{
		return E_FAIL; // ctx が消滅している状態の時は以降で落ちる
	}
	if ( eventname[0] == '\0' ) sprintf(eventname, "%d", (int)dispIdMember);
	sprintf(funcname, "%s%s", m_prefix, eventname);
	if ( m_LogFire ) reportf(m_disps->ctx, "[Event %s]", funcname);
	JSValue global = JS_GetGlobalObject(m_disps->ctx);
	JSValue callfunc = JS_GetPropertyStr(m_disps->ctx, global, funcname);

	if ( !JS_IsUndefined(callfunc) ){
		JSValue argvbuf[MAXVALBUF], *argv = argvbuf;
		if ( pDispParams->cArgs > MAXVALBUF ){
			argv = new JSValue[pDispParams->cArgs];
		}
		for ( unsigned int i = 0 ; i < pDispParams->cArgs; i++ ){
			argv[i] = JS_NewObjectFromVARIANT(m_disps->ctx, &pDispParams->rgvarg[pDispParams->cArgs - 1 - i]);
		}

		JSValue result = JS_Call(m_disps->ctx, callfunc, global, pDispParams->cArgs, argv);
		JS_FreeValue(m_disps->ctx, result); // Event は基本 void なので破棄

		for ( unsigned int i = 0 ; i < pDispParams->cArgs; i++ ){
			JS_FreeValue(m_disps->ctx, argv[i]);
		}
		if ( argv != argvbuf ) delete[] argv;
	}
	JS_FreeValue(m_disps->ctx, callfunc);
	JS_FreeValue(m_disps->ctx, global);
	return S_OK;
}
//-------------------------------------------------- init
void CEventSink::Connect(IConnectionPoint *pPoint)
{
	pPoint->Advise(reinterpret_cast<IUnknown *>(this), &m_ConnectID);
	m_ConnectPoint = pPoint;
}

void ReleaseAllEvent(FREECHAIN *firstchain)
{
	FREECHAIN *chain;

	for(;;){
		CEventSink *pSink;
		chain = firstchain->next;
		if ( chain == NULL ) break;
		pSink = reinterpret_cast<CEventSink *>(chain->value);
		if ( pSink != NULL ) pSink->Disconnect();
	}
}

void CEventSink::Disconnect()
{
	FREECHAIN *chain, *nextchain;
	chain = &(GetCtxInfo(m_disps->ctx)->EventChain);
	for(;;){
		nextchain = chain->next;
		if ( nextchain == NULL ) break;
		if ( nextchain == &(m_disps->chain) ){
			chain->next = nextchain->next;
			break;
		}
		chain = nextchain;
	}

	if ( m_EventType != NULL ){
		if ( m_LogFire ) reportf(m_disps->ctx, "[Disconnect]");
		m_EventType->Release();
		m_EventType = NULL;
	}

	if ( m_ConnectPoint == NULL ) return;
	IConnectionPoint *pConnectPoint = m_ConnectPoint;
	m_ConnectPoint = NULL;
//	reportf(m_disps->ctx, "**DISC**");

	// ↓の中で CEventSink->Release() x3、このため Disconnect の再帰が起きる
	pConnectPoint->Unadvise(m_ConnectID);
	pConnectPoint->Release();


}

void CEventSink::SetEventInfo(ITypeInfo *pEventType)
{
	m_EventType = pEventType;
}

void CEventSink::SetEventTrace(int trace)
{
	m_LogFire = trace;
}

void SetDispGetSet(JSContext *ctx, JSValue &newobj, char *funcnameA, MEMBERID memid, INVOKEKIND invkind)
{
	JSValue getfunc, setfunc;

	if ( invkind & DISPATCH_PROPERTYGET ){
		funcnameA[0] = 'g'; // funcnameA = get_propname
		getfunc = JS_NewCFunctionMagic(ctx,
				reinterpret_cast<JSCFunctionMagic *>(DispatchGetter),
				funcnameA, 0, JS_CFUNC_getter_magic, static_cast<int>(memid));
	}else{
		getfunc = JS_UNDEFINED;
	}

	if ( invkind & DISPATCH_PROPERTYPUT ){
		funcnameA[0] = 's'; // funcnameA = set_propname
		setfunc = JS_NewCFunctionMagic(ctx,
				reinterpret_cast<JSCFunctionMagic *>(DispatchSetter),
				funcnameA, 0, JS_CFUNC_setter_magic, static_cast<int>(memid));
	}else{
		setfunc = JS_UNDEFINED;
	}
	JS_DefinePropertyGetSet(ctx, newobj,
			JS_NewAtom(ctx, funcnameA + FUNCNAMEOFF), getfunc, setfunc, 0);
}

JSValue JS_NewDispatch(JSContext *ctx, IDispatch *pDisp, const char *prefix)
{
	char funcnameA[DISPNAME_MAXLEN + FUNCNAMEOFF + 64];
	JSRuntime *rt = JS_GetRuntime(ctx);

//	reportf(ctx, "[new Dispatch]");
	if ( ClassID_Dispatch == 0 ) JS_NewClassID(&ClassID_Dispatch);
	if ( !JS_IsRegisteredClass(rt, ClassID_Dispatch) ){ // クラスを登録
		JS_NewClass(rt, ClassID_Dispatch, &DispClassDef);
		JSValue ProtoClass = JS_NewObject(ctx);
		JS_SetPropertyFunctionList(ctx, ProtoClass, DispFunctionList, _countof(DispFunctionList));
		JS_SetClassProto(ctx, ClassID_Dispatch, ProtoClass);
	}
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Dispatch);

	DISPATCHSTRUCT *disps = static_cast<DISPATCHSTRUCT *>(malloc(sizeof(DISPATCHSTRUCT)));
	disps->ctx = ctx;
	disps->pDisp = pDisp;
	disps->pEventSink = NULL;
	disps->chain.value = NULL;
	disps->chain.next = NULL;

	JS_SetOpaque(newobj, disps);

	IID eventID;

	if ( prefix != NULL ){ // イベント処理
		InitEventSink(disps, prefix, eventID);
	}

	// メソッド・プロパティをオブジェクトに登録
	strcpy(funcnameA, "set_");

	ITypeInfo *pType;
	ITypeInfo *pEventType = NULL;
	TYPEATTR *attr;

	if ( GetTypeInfoAttr(pDisp, &pType, &attr,
			&pEventType, (disps->pEventSink == NULL) ? NULL : &eventID) ){
		unsigned int funcmax = static_cast<unsigned int>(attr->cFuncs);
		int previd = DISPID_UNKNOWN;
		INVOKEKIND invkind;

		if ( pEventType != NULL ){
			disps->pEventSink->SetEventInfo(pEventType);
		}

		invkind = INVOKE_FUNC;
		for ( unsigned int funcno = 0; funcno < funcmax; funcno++ ){
			FUNCDESC *fdesc;
			BSTR funcnameW;
			unsigned int getnames;

			if ( FAILED(pType->GetFuncDesc(funcno, &fdesc)) ) continue;
			if ( ! (fdesc->wFuncFlags & (FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN)) ){
				if ( invkind != INVOKE_FUNC ){
					if ( fdesc->memid == previd ){ // put & get の場合,まとめる
						SetDispGetSet(ctx, newobj, funcnameA, previd, (INVOKEKIND)(int)((int)fdesc->invkind | (int)invkind));
						invkind = INVOKE_FUNC;
						pType->ReleaseFuncDesc(fdesc);
						continue;
					}
					// 別 ID なので溜めていたプロパティを登録
					SetDispGetSet(ctx, newobj, funcnameA, previd, invkind);
					invkind = INVOKE_FUNC;
				}

				if ( SUCCEEDED(pType->GetNames(fdesc->memid , &funcnameW, 1, &getnames)) ){
					if ( 0 >= WideCharToMultiByte(CP_UTF8, 0, funcnameW, -1, funcnameA + FUNCNAMEOFF, DISPNAME_MAXLEN, NULL, NULL) ){
						strcpy(funcnameA + FUNCNAMEOFF, "?");
					}
					::SysFreeString(funcnameW);
				}else{
					strcpy(funcnameA + FUNCNAMEOFF, "?");
				}

				if ( fdesc->invkind & DISPATCH_METHOD ){ // メソッドはここで登録
					JS_SetPropertyStr(ctx, newobj, funcnameA + FUNCNAMEOFF,
							JS_NewCFunctionMagic(ctx, DispatchFunction,
								funcnameA + FUNCNAMEOFF, 0,
								JS_CFUNC_generic_magic, (int)fdesc->memid));
				}else{
					previd = fdesc->memid;
					if ( fdesc->invkind & (DISPATCH_PROPERTYGET | DISPATCH_PROPERTYPUT) ){ // プロパティはまとめるため登録を次の ID まで保留
						invkind = fdesc->invkind;
					}else{ // プロパティ・メソッド以外はなにもしない
						invkind = INVOKE_FUNC;
					}
				}
			}
			pType->ReleaseFuncDesc(fdesc);
		}
		if ( invkind != INVOKE_FUNC ){ // 遅延していたプロパティを登録
			SetDispGetSet(ctx, newobj, funcnameA, previd, invkind);
		}
		pType->ReleaseTypeAttr(attr);
		pType->Release();
	}
	return newobj;
}

extern "C" JSValue PPxCreateObject(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	const char *str = JS_ToCString(ctx, argv[0]);
	WCHAR progID[MAX_PATH];
	CLSID clsidObject;
	HRESULT hr;

	JSValue oleemulate;
	int emu;

	oleemulate = JS_GetPropertyStr(ctx, this_obj, "emuole");
	JS_ToInt32(ctx, &emu, oleemulate);
	JS_FreeValue(ctx, oleemulate);

	if ( emu ){
		if ( stricmp(str, "scripting.filesystemobject") == 0 ){
			JS_FreeCString(ctx, str);
			return New_FileSystemObject(ctx);
		}else if ( stricmp(str, "adodb.stream") == 0 ){
			JS_FreeCString(ctx, str);
			return New_ADODB_Stream(ctx);
		}
	}

	progID[0] = '\0';
	MultiByteToWideChar(CP_UTF8, 0, str, -1, progID, MAX_PATH);

	if ( SUCCEEDED(hr = ::CLSIDFromProgID(progID, &clsidObject)) ){
		IDispatch *pDisp;

		if ( SUCCEEDED(hr = ::CoCreateInstance(clsidObject, 0,
			CLSCTX_ALL, IID_IDispatch, reinterpret_cast<void**>(&pDisp))) ){

			JS_FreeCString(ctx, str);
			#ifndef RELESE
				CountDisp++;
			#endif
			if ( argc <= 1 ){ // イベント不要
				return JS_NewDispatch(ctx, pDisp, NULL);
			}else{ // イベント通知有り
				const char *prefix = JS_ToCString(ctx, argv[1]);

				JSValue newobj = JS_NewDispatch(ctx, pDisp, prefix);
				JS_FreeCString(ctx, prefix);
				return newobj;
			}
		}
	}

	JS_ThrowReferenceError(ctx, "not support '%s'", str);
	JS_FreeCString(ctx, str);
	return JS_EXCEPTION;
}

extern "C" JSValue PPxGetObject(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR filename[CMDLINESIZE], progidstr[CMDLINESIZE];
	const char *prefix = NULL;
	IDispatch *pDisp = NULL;
	JSValue disp;

	GetJsShortString(ctx, argv[0], filename);
	if ( argc > 1 ){
		GetJsShortString(ctx, argv[1], progidstr);
		if ( argc > 2 ){
			prefix = JS_ToCString(ctx, argv[2]);
		}
	}

	if ( progidstr[0] == '\0' ){
		if ( FAILED(::CoGetObject(filename, 0, IID_IDispatch, reinterpret_cast<void**>(&pDisp))) ){
			pDisp = NULL;
		}
	} else{	// GetInstanceFromFile を使う
		CLSID clsidObject;
		MULTI_QI mq = { &IID_IDispatch, 0, 0 };

		if ( FAILED(::CLSIDFromProgID(progidstr, &clsidObject)) ){
			pDisp = NULL;
		}else{
			if ( SUCCEEDED(::CoGetInstanceFromFile(NULL, &clsidObject,
					NULL, CLSCTX_ALL, STGM_READWRITE, filename, 1, &mq)) ){
				if ( FAILED(mq.pItf->QueryInterface(IID_IDispatch,
						reinterpret_cast<void**>(&pDisp))) ){
					pDisp = NULL;
				}
				mq.pItf->Release();
			}else{
				pDisp = NULL;
			}
		}
	}
	if ( pDisp == NULL ){
		return JS_EXCEPTION;
	}

	disp = JS_NewDispatch(ctx, pDisp, prefix);
	JS_FreeCString(ctx, prefix);
	return disp;
}

extern "C" JSValue PPxConnectObject(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv, int magic)
{
	JSValueConst callargv[2];

	JSValue callfunc = JS_GetPropertyStr(ctx, argv[0], "_");
	if ( !JS_IsObject(callfunc) ) return JS_EXCEPTION;
	callargv[0] = JS_NewString(ctx, magic ? "disconnect" : "connect");
	callargv[1] = argv[1];
	JSValue result = JS_Call(ctx, callfunc, argv[0], 2 - magic, callargv);
	JS_FreeValue(ctx, callargv[0]);
	JS_FreeValue(ctx, callfunc);
	return result;
}
