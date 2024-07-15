/*-----------------------------------------------------------------------------
	Paper Plane xUI QuickJS Script Module							(c)TORO
-----------------------------------------------------------------------------*/
#define STRICT
#define UNICODE
#include <stddef.h>
#include <windows.h>
#include "quickjs.h"
#include "TOROWIN.H"
#include "PPCOMMON.H"
#include "ppxqjs.h"

#define MODIFY_QJS 1 // JS_GetIterator の export patch 済みなら 1

// JSClassID はプロセスグローバル
JSClassID ClassID_PPx = 0;
JSClassID ClassID_Entry = 0;
JSClassID ClassID_Pane = 0;
JSClassID ClassID_Tab = 0;
JSClassID ClassID_Arguments = 0;

DWORD StayInstanseIDserver = ScriptStay_FirstAutoID; // インスタンス番号振出し

typedef union {
	int index;
	WCHAR comment[CMDLINESIZE];
} PPXUPTR_ENTRYCOMMENT;

typedef struct {
	int type;
	WCHAR filename[CMDLINESIZE];
} PPXUPTR_FILEINFORMATION;

typedef struct {
	PPXAPPINFOW *ppxa;
	int PaneIndex;
	int TabIndex;
	BOOL done;
} TABINFO;
#define GetTabInfo(this_obj) ((TABINFO *)JS_GetOpaque(this_obj, ClassID_Tab))

typedef struct {
	PPXAPPINFOW *ppxa;
	int index;
	BOOL done;
} PANEINFO;
#define GetPaneInfo(this_obj) ((PANEINFO *)JS_GetOpaque(this_obj, ClassID_Pane))

//=========================================================== PPx.Tab
JSValue TabItem(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	TABINFO *info = GetTabInfo(this_obj);
	TABINFO *tabinfo = malloc(sizeof(TABINFO));
	*tabinfo = *info;
	if ( argc > 0 ) JS_ToInt32(ctx, &tabinfo->TabIndex, argv[0]);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Tab);
	JS_SetOpaque(newobj, tabinfo);
	return newobj;
}

JSValue GetTabCount(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	int count = info->PaneIndex;

	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABCOUNT, &count);
	return JS_NewInt32(ctx, count);
}

JSValue GetTabIndex(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[2];

	if ( info->TabIndex >= 0 ) return JS_NewInt32(ctx, info->TabIndex);

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABINDEX, &tmp);
	return JS_NewInt32(ctx, tmp[0]);
}

static JSValue SetTabIndex(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	TABINFO *info = GetTabInfo(this_obj);

	JS_ToInt32(ctx, &info->TabIndex, val);
	return JS_UNDEFINED;
}

static JSValue TabIndexFrom(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	TABINFO *info = GetTabInfo(this_obj);
	WCHAR bufw[CMDLINESIZE];
	PPXUPTR_TABINDEXSTRW tmp;
	DWORD_PTR result;

	GetJsShortString(ctx, argv[0], bufw);
	tmp.pane = info->PaneIndex;
	tmp.tab = info->TabIndex;
	tmp.str = bufw;
	result = info->ppxa->Function(info->ppxa, PPXCMDID_COMBOGETTAB, &tmp);
	info->PaneIndex = tmp.pane;
	info->TabIndex = tmp.tab;
	return JS_NewInt32(ctx, (result == PPXA_NO_ERROR));
}

JSValue GetTabPane(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);

	if ( info->PaneIndex >= 0 ){
		return JS_NewInt32(ctx, info->PaneIndex);
	}else {
		int tmp[2];

		tmp[0] = info->PaneIndex;
		tmp[1] = info->TabIndex;
		info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABPANE, &tmp);
		return JS_NewInt32(ctx, tmp[0]);
	}
}

static JSValue SetTabPane(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[4];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	JS_ToInt32(ctx, &tmp[2], val);
	info->PaneIndex = tmp[2];
	info->ppxa->Function(info->ppxa, PPXCMDID_SETCOMBOTABPANE, &tmp);
	return JS_UNDEFINED;
}

JSValue GetTabIDName(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	DWORD buf[REGEXTIDSIZE / (sizeof(DWORD) / sizeof(WCHAR))];

	buf[0] = info->PaneIndex;
	buf[1] = info->TabIndex;
	if ( info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABLONGID, &buf) == 0 ){
		buf[0] = info->PaneIndex;
		buf[1] = info->TabIndex;
		info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABIDNAME, &buf);
		if ( ((WCHAR *)&buf)[1] == '_' ){
			memmove(((WCHAR *)&buf) + 1, ((WCHAR *)&buf) + 2, (strlenW(((WCHAR *)&buf) + 2) + 1) * 2);
		}
	}
	return JS_NewStringW(ctx, (WCHAR *)&buf);
}

JSValue GetTabName(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	DWORD buf[CMDLINESIZE / (sizeof(DWORD) / sizeof(WCHAR))];

	buf[0] = info->PaneIndex;
	buf[1] = info->TabIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABNAME, &buf);
	return JS_NewStringW(ctx, (WCHAR *)&buf);
}

static JSValue SetTabName(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	TABINFO *info = GetTabInfo(this_obj);
	WCHAR bufw[CMDLINESIZE];
	PPXUPTR_TABINDEXSTRW tmp;

	GetJsShortString(ctx, val, bufw);
	tmp.pane = info->PaneIndex;
	tmp.tab = info->TabIndex;
	tmp.str = bufw;
	info->ppxa->Function(info->ppxa, PPXCMDID_SETTABNAME, &tmp);
	return JS_UNDEFINED;
}

JSValue GetTabType(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[2];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOWNDTYPE, &tmp);
	return JS_NewInt32(ctx, tmp[0]);
}

static JSValue SetTabType(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[4];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	JS_ToInt32(ctx, &tmp[2], val);
	info->ppxa->Function(info->ppxa, PPXCMDID_SETCOMBOWNDTYPE, &tmp);
	return JS_UNDEFINED;
}

JSValue GetTabLock(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[2];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_DIRLOCK, &tmp);
	return JS_NewInt32(ctx, tmp[0]);
}

static JSValue SetTabLock(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[4];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	JS_ToInt32(ctx, &tmp[2], val);
	info->ppxa->Function(info->ppxa, PPXCMDID_SETDIRLOCK, &tmp);
	return JS_UNDEFINED;
}

JSValue GetTabColor(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[2];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_TABTEXTCOLOR, &tmp);
	return JS_NewInt32(ctx, tmp[0]);
}

static JSValue SetTabColor(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[4];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	JS_ToInt32(ctx, &tmp[2], val);
	info->ppxa->Function(info->ppxa, PPXCMDID_SETTABTEXTCOLOR, &tmp);
	return JS_UNDEFINED;
}

JSValue GetTabBackColor(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[2];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_TABBACKCOLOR, &tmp);
	return JS_NewInt32(ctx, tmp[0]);
}

static JSValue SetTabBackColor(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	TABINFO *info = GetTabInfo(this_obj);
	int tmp[4];

	tmp[0] = info->PaneIndex;
	tmp[1] = info->TabIndex;
	JS_ToInt32(ctx, &tmp[2], val);
	info->ppxa->Function(info->ppxa, PPXCMDID_SETTABBACKCOLOR, &tmp);
	return JS_UNDEFINED;
}

JSValue TabExecute(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	TABINFO *info = GetTabInfo(this_obj);
	WCHAR parambuf[CMDLINESIZE], *param;
	PPXUPTR_TABINDEXSTRW tmp;

	int resultcode;

	param = GetJsLongString(ctx, argv[0], parambuf);
	tmp.pane = info->PaneIndex;
	tmp.tab = info->TabIndex;
	tmp.str = param;

	resultcode = (int)info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABEXECUTE, &tmp);
	FreeJsLongString(param, parambuf);
	if ( resultcode == 1 ) resultcode = NO_ERROR;
	return JS_NewInt32(ctx, resultcode);
}

JSValue TabExtract(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	TABINFO *info = GetTabInfo(this_obj);
	WCHAR parambuf[CMDLINESIZE], *param;
	PPXUPTR_TABINDEXSTRW tmp;
	LONG_PTR long_result;

	param = GetJsLongString(ctx, argv[0], parambuf);
	tmp.pane = info->PaneIndex;
	tmp.tab = info->TabIndex;
	tmp.str = param;

	long_result = (LONG_PTR)info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABEXTRACTLONG, &tmp);

	if ( long_result >= 0x10000 ){
		BSTR b_result;

		FreeJsLongString(param, parambuf);
		b_result = (BSTR)long_result;
		if (b_result == NULL){
			return JS_NewString(ctx, "");
		}else {
			JSValue result;

			result = JS_NewStringW(ctx, b_result);
			SysFreeString(b_result);
			return result;
		}
	}

	tmp.pane = info->PaneIndex;
	tmp.tab = info->TabIndex;
	tmp.str = param;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABEXTRACT, &tmp);
	FreeJsLongString(param, parambuf);
	return JS_NewStringW(ctx, parambuf);
}

JSValue TabatEnd(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	TABINFO *info = GetTabInfo(this_obj);
	int maxtabs = 0;

	if ( info->TabIndex < 0) info->TabIndex = 0;
	maxtabs = info->PaneIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABCOUNT, &maxtabs);
	return JS_NewInt32(ctx, (info->TabIndex >= maxtabs) ? 1 : 0);
}

JSValue TabmoveNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	GetTabInfo(this_obj)->TabIndex++;
	return JS_UNDEFINED;
}

JSValue TabReset(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	GetTabInfo(this_obj)->TabIndex = -1;
	return JS_UNDEFINED;
}

JSValue GetTabIteratorValue(JSContext *ctx, JSValueConst this_obj)
{
	return JS_DupValue(ctx, this_obj);
}

JSValue GetTabIteratorDone(JSContext *ctx, JSValueConst this_obj)
{
	TABINFO *info = GetTabInfo(this_obj);

	return JS_NewBool(ctx, info->done);
}

JSValue TabIteratorNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	TABINFO *info = GetTabInfo(this_obj);

	int maxtabs = 0;

	if ( info->TabIndex < 0) info->TabIndex = 0;

	maxtabs = info->PaneIndex;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOTABCOUNT, &maxtabs);
	info->done = info->TabIndex >= maxtabs;
	return JS_DupValue(ctx, this_obj);
}

JSValue TabIterator(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	TABINFO *info = GetTabInfo(this_obj);

	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Pane);
	TABINFO *tabinfo = malloc(sizeof(TABINFO));

	tabinfo->ppxa = info->ppxa;
	tabinfo->PaneIndex = info->PaneIndex;
	tabinfo->TabIndex = -1;
	tabinfo->done = TRUE;

	JS_SetOpaque(newobj, tabinfo);
	return newobj;
}

void Tabfinalizer(JSRuntime *rt, JSValue val)
{
	free(JS_GetOpaque(val, ClassID_Tab));
}

static const JSCFunctionListEntry TabFunctionList[] = {
	JS_CFUNC_DEF("Item", 0, TabItem),
	JS_CFUNC_DEF("item", 0, TabItem),
	JS_CGETSET_DEF("Count", GetTabCount, NULL),
	JS_CGETSET_DEF("length", GetTabCount, NULL),
	JS_CGETSET_DEF("Index", GetTabIndex, SetTabIndex),
	JS_CGETSET_DEF("index", GetTabIndex, SetTabIndex),
	JS_CFUNC_DEF("IndexFrom", 1, TabIndexFrom),
	JS_CGETSET_DEF("Pane", GetTabPane, SetTabPane),
	JS_CGETSET_DEF("IDName", GetTabIDName, NULL),
	JS_CGETSET_DEF("Name", GetTabName, SetTabName),
	JS_CGETSET_DEF("Type", GetTabType, SetTabType),
	JS_CGETSET_DEF("Lock", GetTabLock, SetTabLock),
	JS_CGETSET_DEF("Color", GetTabColor, SetTabColor),
	JS_CGETSET_DEF("BackColor", GetTabBackColor, SetTabBackColor),
	JS_CFUNC_DEF("Execute", 0, TabExecute),
	JS_CFUNC_DEF("Extract", 0, TabExtract),
	JS_CFUNC_DEF("atEnd", 0, TabatEnd),
	JS_CFUNC_DEF("moveNext", 0, TabmoveNext),
	JS_CFUNC_DEF("Reset", 0, TabReset),
#if !MODIFY_QJS
	JS_CFUNC_DEF("Symbol.iterator", 0, TabIterator),
#endif
	JS_CFUNC_DEF("[Symbol.iterator]", 0, TabIterator),
	JS_CFUNC_DEF("next", 0, TabIteratorNext),
	JS_CGETSET_DEF("value", GetTabIteratorValue, NULL),
	JS_CGETSET_DEF("done", GetTabIteratorDone, NULL),
};

JSClassDef TabClassDef = {
	"_TaB_", // class_name
	Tabfinalizer, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};

static JSValue GetPaneTab(JSContext *ctx, JSValueConst this_obj)
{
	PANEINFO *info = GetPaneInfo(this_obj);
	JSRuntime *rt = JS_GetRuntime(ctx);

	if ( ClassID_Tab == 0 ) JS_NewClassID(&ClassID_Tab);
	if ( !JS_IsRegisteredClass(rt, ClassID_Tab) ){ // クラスを登録
		JS_NewClass(rt, ClassID_Tab, &TabClassDef);
		JSValue ProtoClass = JS_NewObject(ctx);
		JS_SetPropertyFunctionList(ctx, ProtoClass, TabFunctionList, _countof(TabFunctionList));
		JS_SetClassProto(ctx, ClassID_Tab, ProtoClass);
	}
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Tab);
	TABINFO *tabinfo = malloc(sizeof(TABINFO));
	tabinfo->ppxa = info->ppxa;
	tabinfo->PaneIndex = info->index;
	tabinfo->TabIndex = -1;
	tabinfo->done = TRUE;
	JS_SetOpaque(newobj, tabinfo);
	return newobj;
}

//=========================================================== PPx.Pane

JSValue PaneItem(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PANEINFO *info = GetPaneInfo(this_obj);
	PANEINFO *paneinfo = malloc(sizeof(PANEINFO));
	*paneinfo = *info;
	if ( argc > 0 ) JS_ToInt32(ctx, &paneinfo->index, argv[0]);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Pane);
	JS_SetOpaque(newobj, paneinfo);
	return newobj;
}

JSValue GetPaneCount(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetPaneInfo(this_obj)->ppxa;
	int count;

	ppxa->Function(ppxa, PPXCMDID_COMBOSHOWPANES, &count);
	return JS_NewInt32(ctx, count);
}

JSValue GetPaneIndex(JSContext *ctx, JSValueConst this_obj)
{
	PANEINFO *info = GetPaneInfo(this_obj);
	int tmp[2];

	if ( info->index >= 0 ) return JS_NewInt32(ctx, info->index);

	tmp[0] = info->index;
	tmp[1] = 0; // 未使用であるが、念のため
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOSHOWINDEX, &tmp);
	return JS_NewInt32(ctx, tmp[0]);
}

static JSValue SetPaneIndex(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PANEINFO *info = GetPaneInfo(this_obj);

	JS_ToInt32(ctx, &info->index, val);
	return JS_UNDEFINED;
}

JSValue GetPaneGroupIndex(JSContext *ctx, JSValueConst this_obj)
{
	PANEINFO *info = GetPaneInfo(this_obj);
	int tmp[2];

	tmp[0] = info->index;
	tmp[1] = -1;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOGROUPINDEX, &tmp);
	return JS_NewInt32(ctx, tmp[1]);
}

static JSValue SetPaneGroupIndex(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PANEINFO *info = GetPaneInfo(this_obj);

	int tmp[2];

	tmp[0] = info->index;
	JS_ToInt32(ctx, &tmp[1], val);
	info->ppxa->Function(info->ppxa, PPXCMDID_SETCOMBOGROUPINDEX, &tmp);
	return JS_UNDEFINED;
}

JSValue GetPaneGroupCount(JSContext *ctx, JSValueConst this_obj)
{
	PANEINFO *info = GetPaneInfo(this_obj);
	int tmp[2];

	tmp[0] = info->index;
	tmp[1] = 0;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOGROUPCOUNT, &tmp);
	return JS_NewInt32(ctx, tmp[1]);
}

JSValue GetPaneGroupList(JSContext *ctx, JSValueConst this_obj)
{
	int tmpc[2];
	PPXUPTR_TABINDEXSTRW tmp;
	PANEINFO *info = GetPaneInfo(this_obj);
	WCHAR param[CMDLINESIZE];
	int index = 0, maxindex;

	tmpc[0] = info->index;
	tmpc[1] = 0;
	if ( info->ppxa->Function(info->ppxa, PPXCMDID_COMBOGROUPCOUNT, &tmpc) != PPXA_NO_ERROR ){
		return JS_NULL;
	}
	maxindex = tmpc[1];

	JSValue array = JS_NewArray(ctx);

	for (; index < maxindex; index++){
		JSValue jsitem;

		tmp.pane = info->index;
		tmp.tab = index;
		tmp.str = param;
		param[0] = '\0';
		if ( info->ppxa->Function(info->ppxa, PPXCMDID_COMBOGROUPNAME, &tmp) !=  PPXA_NO_ERROR ){
			break;
		}
		jsitem = JS_NewStringW(ctx, param);
		JS_SetPropertyUint32(ctx, array, index, jsitem);
	}
	return array;
}

JSValue GetPaneGroupName(JSContext *ctx, JSValueConst this_obj)
{
	PPXUPTR_TABINDEXSTRW tmp;
	PANEINFO *info = GetPaneInfo(this_obj);
	WCHAR param[CMDLINESIZE];

	tmp.pane = info->index;
	tmp.tab = -1;
	tmp.str = param;
	param[0] = '\0';
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOGROUPNAME, &tmp);
	return JS_NewStringW(ctx, param);
}

static JSValue SetPaneGroupName(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXUPTR_TABINDEXSTRW tmp;
	PANEINFO *info = GetPaneInfo(this_obj);
	WCHAR param[CMDLINESIZE];

	tmp.pane = info->index;
	tmp.tab = -1;
	tmp.str = param;
	GetJsShortString(ctx, val, param);
	info->ppxa->Function(info->ppxa, PPXCMDID_SETCOMBOGROUPNAME, &tmp);
	return JS_UNDEFINED;
}

static JSValue PaneIndexFrom(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PANEINFO *info = GetPaneInfo(this_obj);
	WCHAR bufw[CMDLINESIZE];

	GetJsShortString(ctx, argv[0], bufw);

	info->index = (int)info->ppxa->Function(info->ppxa,
			PPXCMDID_COMBOGETPANE, (void *)bufw);
	return JS_NewInt32(ctx, (info->index >= 0));
}

JSValue PaneatEnd(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PANEINFO *info = GetPaneInfo(this_obj);
	int maxpanes = 0;

	if ( info->index < 0) info->index = 0;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOSHOWPANES, &maxpanes);
	return JS_NewInt32(ctx, (info->index >= maxpanes) ? 1 : 0);
}

JSValue PanemoveNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PANEINFO *info = GetPaneInfo(this_obj);

	info->index++;
	return JS_UNDEFINED;
}

JSValue PaneReset(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PANEINFO *info = GetPaneInfo(this_obj);

	info->index = -1;
	return JS_UNDEFINED;
}

JSValue GetPaneIteratorValue(JSContext *ctx, JSValueConst this_obj)
{
	return JS_DupValue(ctx, this_obj);
}

JSValue GetPaneIteratorDone(JSContext *ctx, JSValueConst this_obj)
{
	PANEINFO *info = GetPaneInfo(this_obj);

	return JS_NewBool(ctx, info->done);
}

JSValue PaneIteratorNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PANEINFO *info = GetPaneInfo(this_obj);

	int maxpanes = 0;

	if ( info->index < 0) info->index = 0;
	info->ppxa->Function(info->ppxa, PPXCMDID_COMBOSHOWPANES, &maxpanes);
	info->done = info->index >= maxpanes;
	return JS_DupValue(ctx, this_obj);
}

JSValue PaneIterator(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PANEINFO *info = GetPaneInfo(this_obj);

	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Pane);
	PANEINFO *paneinfo = malloc(sizeof(PANEINFO));
	paneinfo->ppxa = info->ppxa;
	paneinfo->index = -1;
	paneinfo->done = TRUE;
	JS_SetOpaque(newobj, paneinfo);
	return newobj;
}

static const JSCFunctionListEntry PaneFunctionList[] = {
	JS_CFUNC_DEF("Item", 0, PaneItem),
	JS_CFUNC_DEF("item", 0, PaneItem),
	JS_CGETSET_DEF("Count", GetPaneCount, NULL),
	JS_CGETSET_DEF("length", GetPaneCount, NULL),
	JS_CGETSET_DEF("Index", GetPaneIndex, SetPaneIndex),
	JS_CGETSET_DEF("index", GetPaneIndex, SetPaneIndex),
	JS_CFUNC_DEF("IndexFrom", 1, PaneIndexFrom),
	JS_CGETSET_DEF("Tab", GetPaneTab, NULL),
	JS_CFUNC_DEF("atEnd", 0, PaneatEnd),
	JS_CFUNC_DEF("moveNext", 0, PanemoveNext),
	JS_CFUNC_DEF("Reset", 0, PaneReset),

	JS_CGETSET_DEF("GroupIndex", GetPaneGroupIndex, SetPaneGroupIndex),
	JS_CGETSET_DEF("GroupCount", GetPaneGroupCount, NULL),
	JS_CGETSET_DEF("GroupList", GetPaneGroupList, NULL),
	JS_CGETSET_DEF("GroupName", GetPaneGroupName, SetPaneGroupName),

#if !MODIFY_QJS
	JS_CFUNC_DEF("Symbol.iterator", 0, PaneIterator),
#endif
	JS_CFUNC_DEF("[Symbol.iterator]", 0, PaneIterator),
	JS_CFUNC_DEF("next", 0, PaneIteratorNext),
	JS_CGETSET_DEF("value", GetPaneIteratorValue, NULL),
	JS_CGETSET_DEF("done", GetPaneIteratorDone, NULL),
};

void Panefinalizer(JSRuntime *rt, JSValue val)
{
	free(JS_GetOpaque(val, ClassID_Pane));
}

JSClassDef PaneClassDef = {
	"_PanE_", // class_name
	Panefinalizer, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};

static JSValue PPxGetPane(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	JSRuntime *rt = JS_GetRuntime(ctx);

	if ( ClassID_Pane == 0 ) JS_NewClassID(&ClassID_Pane);
	if ( !JS_IsRegisteredClass(rt, ClassID_Pane) ){ // クラスを登録
		JS_NewClass(rt, ClassID_Pane, &PaneClassDef);
		JSValue ProtoClass = JS_NewObject(ctx);
		JS_SetPropertyFunctionList(ctx, ProtoClass, PaneFunctionList, _countof(PaneFunctionList));
		JS_SetClassProto(ctx, ClassID_Pane, ProtoClass);
	}
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Pane);
	PANEINFO *paneinfo = malloc(sizeof(PANEINFO));
	paneinfo->ppxa = ppxa;
	paneinfo->index = -1;
	paneinfo->done = TRUE;
	JS_SetOpaque(newobj, paneinfo);
	return newobj;
}

//=========================================================== PPx.Entry
#define ENUMENTRY_WITHMARK 1 // マーク順に列挙 マーク無しの時はカーソル位置のみ
#define ENUMENTRY_MARKONLY 2 // マーク順に列挙 マーク無しの時は列挙無し
#define ENUMENTRY_INDEX 3 // index 順に列挙
typedef struct {
	PPXAPPINFOW *ppxa;
	int index;
	int enummode; // 正:列挙未開始、負:列挙中
	BOOL done;
} ENTRYINFO;

#define GetEntryInfo(this_obj) ((ENTRYINFO *)JS_GetOpaque(this_obj, ClassID_Entry))

JSValue EntryItem(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Entry);
	ENTRYINFO *newinfo = malloc(sizeof(ENTRYINFO));
	newinfo->ppxa = info->ppxa;
	newinfo->index = info->index;
	newinfo->enummode = ENUMENTRY_WITHMARK;
	newinfo->done = TRUE;
	if ( argc > 0 ) JS_ToInt32(ctx, &newinfo->index, argv[0]);
	if ( newinfo->index == -1 ){
		info->ppxa->Function(info->ppxa, PPXCMDID_CSRINDEX, &newinfo->index);
	}
	JS_SetOpaque(newobj, newinfo);
	return newobj;
}

JSValue GetEntryCount(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetEntryInfo(this_obj)->ppxa;
	int count;

	ppxa->Function(ppxa, PPXCMDID_DIRTTOTAL, &count);
	return JS_NewInt32(ctx, count);
}

JSValue GetEntryIndex(JSContext *ctx, JSValueConst this_obj)
{
	return JS_NewInt32(ctx, GetEntryInfo(this_obj)->index);
}

static JSValue SetEntryIndex(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int index, maxindex;

	JS_ToInt32(ctx, &index, val);
	if ( index >= 0 ){
		info->ppxa->Function(info->ppxa, PPXCMDID_DIRTTOTAL, &maxindex);
		if ( index < maxindex ) info->index = index;
	}else{
		info->ppxa->Function(info->ppxa, PPXCMDID_CSRINDEX, &info->index);
	}
	return JS_UNDEFINED;
}

JSValue GetEntryMark(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int Value;

	Value = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYMARK, &Value);
	return JS_NewInt32(ctx, Value);
}

static JSValue SetEntryMark(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int nums[2];

	JS_ToInt32(ctx, &nums[0], val);
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSETMARK, nums);
	info->index = nums[1]; // 加工済みの値を回収して高速化
	return JS_UNDEFINED;
}

JSValue GetEntryPointMark(JSContext *ctx, JSValueConst this_obj, int magic)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int result;

	result = info->index;
	info->ppxa->Function(info->ppxa, magic, &result);
	if ( result == -1 ){
		result = 0;
	}else {
		info->index = result;
		result = 1;
	}
	return JS_NewInt32(ctx, result);
}

JSValue GetEntryName(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	PPXUPTR_INDEX_UPATHW tmp;

	tmp.index = info->index;
	if ( info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYNAME, &tmp) == PPXA_NO_ERROR ){
		return JS_NewStringW(ctx, tmp.path);
	}else{
		JS_ThrowReferenceError(ctx, "entry not support");
		return JS_EXCEPTION;
	}
}

JSValue GetEntryShortName(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	PPXUPTR_INDEX_UPATHW tmp;

	tmp.index = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYANAME, &tmp);
	if ( tmp.path[0] == '\0' ){
		tmp.index = info->index;
		info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYNAME, tmp.path);
	}
	return JS_NewStringW(ctx, tmp.path);
}

JSValue GetEntryAttributes(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int value;

	value = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYATTR, &value);
	return JS_NewInt32(ctx, value);
}

JSValue GetEntrySize(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	__int64 value;

	value = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYMSIZE, &value);
	return JS_NewInt64(ctx, value);
}

const int EntruTimeType[3] = {PPXCMDID_ENTRYCTIME, PPXCMDID_ENTRYATIME, PPXCMDID_ENTRYMTIME};
JSValue GetEntryTime(JSContext *ctx, JSValueConst this_obj, int magic)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	__int64 utctime;

	utctime = info->index;
	info->ppxa->Function(info->ppxa, EntruTimeType[magic], &utctime);
	return JS_NewInt64(ctx, (utctime - 116444736000000000) / 10000);
}

JSValue GetEntryState(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int value;

	value = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSTATE, &value);
	return JS_NewInt32(ctx, value & 0x1f);
}

static JSValue SetEntryState(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int nums[2];
	int value = 0;

	JS_ToInt32(ctx, &value, val);

	nums[0] = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSTATE, nums);

	nums[0] = (nums[0] & 0xffe0) | value;
	nums[1] = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSETSTATE, nums);
	info->index = nums[1]; // 加工済みの値を回収して高速化
	return JS_UNDEFINED;
}

JSValue GetEntryExtColor(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int value;

	value = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYEXTCOLOR, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue SetEntryExtColor(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int nums[2];
	int value = 0;

	JS_ToInt32(ctx, &value, val);

	nums[0] = value;
	nums[1] = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSETEXTCOLOR, nums);
	info->index = nums[1]; // 加工済みの値を回収して高速化
	return JS_UNDEFINED;
}

static JSValue GetEntryComment(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	PPXUPTR_ENTRYCOMMENT tmp;

	tmp.index = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYCOMMENT, &tmp);
	return JS_NewStringW(ctx, tmp.comment);
}

static JSValue SetEntryComment(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	PPXUPTR_ENTRYCOMMENT tmp;

	tmp.index = info->index;
	GetJsShortString(ctx, val, tmp.comment);
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSETCOMMENT, &tmp);
	return JS_UNDEFINED;
}

JSValue EntryGetComment(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	ENTRYEXTDATASTRUCT eeds;
	WCHAR bufw[CMDLINESIZE];
	int id = 0;

	JS_ToInt32(ctx, &id, argv[0]);
	eeds.entry = info->index;
	eeds.id = (id > 100) ? id : (DFC_COMMENTEX_MAX - (id - 1));
	eeds.size = sizeof(bufw);
	eeds.data = (BYTE *)bufw;
	if ( info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYEXTDATA_GETDATA, &eeds) == 0 ){
		bufw[0] = '\0';
	}
	return JS_NewStringW(ctx, bufw);
}

JSValue EntrySetComment(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	ENTRYEXTDATASTRUCT eeds;
	WCHAR commentbuf[CMDLINESIZE], *comment;
	int id = 0;

	JS_ToInt32(ctx, &id, argv[0]);
	comment = GetJsLongString(ctx, argv[1], commentbuf);

	eeds.entry = info->index;
	eeds.id = (id > 100) ? id : (DFC_COMMENTEX_MAX - (id - 1));
	eeds.size = (wcslen(comment) + 1) * sizeof(WCHAR);
	eeds.data = (BYTE *)comment;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYEXTDATA_SETDATA, &eeds);
	FreeJsLongString(comment, commentbuf);
	return JS_UNDEFINED;
}

JSValue GetEntryInformation(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	PPXUPTR_ENTRYINFOW entryinfo;
	JSValue result;

	entryinfo.index = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYINFO, &entryinfo);
	if ( (DWORD_PTR)entryinfo.index < 0x10000 ) return JS_NULL;
	result = JS_NewStringW(ctx, entryinfo.result);
	HeapFree(ProcHeap, 0, (void *)entryinfo.result);
	return result;
}

JSValue GetEntryHide(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int value;

	value = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYHIDEENTRY, &value);
	return JS_NewInt32(ctx, info->index);
}

JSValue GetEntryHighlight(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int value;

	value = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSTATE, &value);
	return JS_NewInt32(ctx, value >> 5);
}

static JSValue SetEntryHighlight(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	int nums[2];
	int value = 0;

	JS_ToInt32(ctx, &value, val);

	nums[0] = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSTATE, nums);

	nums[0] = (nums[0] & 0x1f) | (value << 5);
	nums[1] = info->index;
	info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYSETSTATE, nums);
	info->index = nums[1]; // 加工済みの値を回収して高速化
	return JS_UNDEFINED;
}

BOOL IsEntryVailed(PPXAPPINFOW *m_pxa, int newindex)
{
	int state;
	PPXUPTR_INDEX_UPATHW tmp;

	state = newindex;
	m_pxa->Function(m_pxa, PPXCMDID_ENTRYSTATE, &state);
	if ( (state & 7) < 2 ) return FALSE; // < ECS_NORMAL

	tmp.index = newindex;
	m_pxa->Function(m_pxa, PPXCMDID_ENTRYNAME, &tmp);
	if ( (tmp.path[0] == '.') &&
		 ((tmp.path[1] == '\0') ||
		  ((tmp.path[1] == '.') && (tmp.path[2] == '\0') ) ) ){ // . or ..
		return FALSE;
	}
	return TRUE;
}

JSValue EntryReset(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);

	info->index = -1;
	info->enummode = ENUMENTRY_WITHMARK;
	return JS_UNDEFINED;
}

BOOL CheckNextEntryAt(ENTRYINFO *info, int newindex) // ENUMENTRY_INDEX
{
	int maxindex;

	info->ppxa->Function(info->ppxa, PPXCMDID_DIRTTOTAL, &maxindex);
	for (;;) {
		if ( newindex >= maxindex) return TRUE;
		if ( IsEntryVailed(info->ppxa, newindex) ) {
			info->index = newindex;
			return FALSE;
		}
		newindex++;
	}
}

BOOL EntryCheckAtEnd(ENTRYINFO *info)
{
	int result;

	if ( info->enummode > 0 ){ // 初めて…列挙初期化
		info->enummode = -info->enummode;
		if ( info->enummode == -ENUMENTRY_INDEX ){
			return CheckNextEntryAt(info, 0);
		}

		info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYMARKFIRST_HS, &result);
		if ( result == -1 ){ // マーク無し
			if ( info->enummode == -ENUMENTRY_WITHMARK ){
				info->ppxa->Function(info->ppxa, PPXCMDID_CSRINDEX, &result);
				info->index = result;
			}else { // ENUMENTRY_MARKONLY
				return TRUE;
			}
		}else {
			info->index = result;
		}
		if ( IsEntryVailed(info->ppxa, info->index) ) return FALSE;
	}

	// ２回目以降
	if ( info->enummode == -ENUMENTRY_INDEX ){
		return CheckNextEntryAt(info, info->index + 1);
	}
	// ENUMENTRY_WITHMARK / ENUMENTRY_MARKONLY
	for (;;){
		result = info->index;
		info->ppxa->Function(info->ppxa, PPXCMDID_ENTRYMARKNEXT_HS, &result);
		if ( (result == -1) || (result == info->index)) return TRUE;
		info->index = result;
		if ( IsEntryVailed(info->ppxa, result) ) return FALSE;
	}
}

JSValue EntryatEnd(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return JS_NewBool(ctx, EntryCheckAtEnd(GetEntryInfo(this_obj)));
}

JSValue EntrymoveNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return JS_UNDEFINED;
}

JSValue GetEntryAllMark(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Entry);
	ENTRYINFO *it_info = malloc(sizeof(ENTRYINFO));
	it_info->ppxa = info->ppxa;
	it_info->index = 0;
	it_info->enummode = ENUMENTRY_MARKONLY;
	it_info->done = FALSE;
	JS_SetOpaque(newobj, it_info);
	return newobj;
}

JSValue GetEntryAllEntry(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Entry);
	ENTRYINFO *it_info = malloc(sizeof(ENTRYINFO));
	it_info->ppxa = info->ppxa;
	it_info->index = 0;
	it_info->enummode = ENUMENTRY_INDEX;
	it_info->done = FALSE;
	JS_SetOpaque(newobj, it_info);
	return newobj;
}

JSValue GetEntryIteratorValue(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);

	if ( info->enummode >= 0 ) return JS_UNDEFINED;
	return JS_DupValue(ctx, this_obj);
}

JSValue GetEntryIteratorDone(JSContext *ctx, JSValueConst this_obj)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);

	if ( info->enummode >= 0 ) return JS_UNDEFINED;
	return JS_NewBool(ctx, info->done);
}

JSValue EntryIteratorNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);

	info->done = EntryCheckAtEnd(info);
	return JS_DupValue(ctx, this_obj);
}

JSValue EntryIterator(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ENTRYINFO *info = GetEntryInfo(this_obj);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Entry);
	ENTRYINFO *it_info = malloc(sizeof(ENTRYINFO));
	it_info->ppxa = info->ppxa;
	it_info->index = 0;
	it_info->enummode = (info->enummode > 0) ? info->enummode : -info->enummode;
	it_info->done = FALSE;
	JS_SetOpaque(newobj, it_info);
	return newobj;
}

static const JSCFunctionListEntry EntryFunctionList[] = {
	JS_CFUNC_DEF("Item", 0, EntryItem),
	JS_CFUNC_DEF("item", 0, EntryItem),
	JS_CGETSET_DEF("Count", GetEntryCount, NULL),
	JS_CGETSET_DEF("length", GetEntryCount, NULL),
	JS_CGETSET_DEF("Index", GetEntryIndex, SetEntryIndex),
	JS_CGETSET_DEF("index", GetEntryIndex, SetEntryIndex),
	JS_CGETSET_DEF("Mark", GetEntryMark, SetEntryMark),
	JS_CGETSET_MAGIC_DEF("FirstMark", GetEntryPointMark, NULL, PPXCMDID_ENTRYMARKFIRST_HS),
	JS_CGETSET_MAGIC_DEF("NextMark", GetEntryPointMark, NULL, PPXCMDID_ENTRYMARKNEXT_HS),
	JS_CGETSET_MAGIC_DEF("PrevMark", GetEntryPointMark, NULL, PPXCMDID_ENTRYMARKPREV_HS),
	JS_CGETSET_MAGIC_DEF("LastMark", GetEntryPointMark, NULL, PPXCMDID_ENTRYMARKLAST_HS),
	JS_CGETSET_DEF("Name", GetEntryName, NULL),
	JS_CGETSET_DEF("ShortName", GetEntryShortName, NULL),
	JS_CGETSET_DEF("Attributes", GetEntryAttributes, NULL),
	JS_CGETSET_DEF("Size", GetEntrySize, NULL),
	JS_CGETSET_MAGIC_DEF("DateCreated", GetEntryTime, NULL, 0),
	JS_CGETSET_MAGIC_DEF("DateLastAccessed", GetEntryTime, NULL, 1),
	JS_CGETSET_MAGIC_DEF("DateLastModified", GetEntryTime, NULL, 2),
	JS_CGETSET_DEF("State", GetEntryState, SetEntryState),
	JS_CGETSET_DEF("ExtColor", GetEntryExtColor, SetEntryExtColor),
	JS_CGETSET_DEF("Comment", GetEntryComment, SetEntryComment),
	JS_CFUNC_DEF("GetComment", 1, EntryGetComment),
	JS_CFUNC_DEF("SetComment", 2, EntrySetComment),
	JS_CGETSET_DEF("Information", GetEntryInformation, NULL),
	JS_CGETSET_DEF("Hide", GetEntryHide, NULL),
	JS_CGETSET_DEF("Highlight", GetEntryHighlight, SetEntryHighlight),
	JS_CGETSET_DEF("AllMark", GetEntryAllMark, NULL),
	JS_CGETSET_DEF("AllEntry", GetEntryAllEntry, NULL),
	JS_CFUNC_DEF("atEnd", 0, EntryatEnd),
	JS_CFUNC_DEF("moveNext", 0, EntrymoveNext),
	JS_CFUNC_DEF("Reset", 0, EntryReset),
#if !MODIFY_QJS
	JS_CFUNC_DEF("Symbol.iterator", 0, EntryIterator),
#endif
	JS_CFUNC_DEF("[Symbol.iterator]", 0, EntryIterator),
	JS_CFUNC_DEF("next", 0, EntryIteratorNext),
	JS_CGETSET_DEF("value", GetEntryIteratorValue, NULL),
	JS_CGETSET_DEF("done", GetEntryIteratorDone, NULL),
};

void Entryfinalizer(JSRuntime *rt, JSValue val)
{
	free(JS_GetOpaque(val, ClassID_Entry));
}

JSClassDef EntryClassDef = {
	"_EntrY_", // class_name
	Entryfinalizer, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};

static JSValue PPxGetEntry(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	JSRuntime *rt = JS_GetRuntime(ctx);

	if ( ClassID_Entry == 0 ) JS_NewClassID(&ClassID_Entry);
	if ( !JS_IsRegisteredClass(rt, ClassID_Entry) ){ // クラスを登録
		JS_NewClass(rt, ClassID_Entry, &EntryClassDef);
		JSValue ProtoClass = JS_NewObject(ctx);
		JS_SetPropertyFunctionList(ctx, ProtoClass, EntryFunctionList, _countof(EntryFunctionList));
		JS_SetClassProto(ctx, ClassID_Entry, ProtoClass);
	}
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Entry);
	ENTRYINFO *entryinfo = malloc(sizeof(ENTRYINFO));
	entryinfo->ppxa = ppxa;
	entryinfo->index = 0;
	entryinfo->enummode = ENUMENTRY_WITHMARK;
	entryinfo->done = FALSE;
	ppxa->Function(ppxa, PPXCMDID_CSRINDEX, &entryinfo->index);
	JS_SetOpaque(newobj, entryinfo);
	return newobj;
}

//=========================================================== PPx.Arguments
typedef struct {
	PPXMCOMMANDSTRUCT *pxc;
	int index;
} ARGUMENTSINFO;
#define GetArgumentsInfo(this_obj) ((ARGUMENTSINFO *)JS_GetOpaque(this_obj, ClassID_Arguments))

JSValue ArgumentsItem(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ARGUMENTSINFO *info = GetArgumentsInfo(this_obj);
	ARGUMENTSINFO *newinfo = malloc(sizeof(ARGUMENTSINFO));
	*newinfo = *info;
	if ( argc > 0 ) JS_ToInt32(ctx, &newinfo->index, argv[0]);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Arguments);
	JS_SetOpaque(newobj, newinfo);
	return newobj;
}

JSValue GetArgumentsCount(JSContext *ctx, JSValueConst this_obj)
{
	return JS_NewInt32(ctx, GetArgumentsInfo(this_obj)->pxc->paramcount - 1);
}

JSValue GetArgumentsIndex(JSContext *ctx, JSValueConst this_obj)
{
	return JS_NewInt32(ctx, GetArgumentsInfo(this_obj)->index);
}

JSValue GetArgumentsValue(JSContext *ctx, JSValueConst this_obj)
{
	ARGUMENTSINFO *info = GetArgumentsInfo(this_obj);

	int index = info->index;

	if ( (index >= 0) && (index < (info->pxc->paramcount - 1)) ){
		const WCHAR *argptr;

		argptr = info->pxc->param;
		for ( ;;){
			argptr += wcslen(argptr) + 1;
			if ( index-- <= 0 ) break;
		}
		return JS_NewStringW(ctx, argptr);
	}
	return JS_NULL;
}

JSValue ArgumentsToString(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return GetArgumentsValue(ctx, this_obj);
}

static JSValue SetArgumentsIndex(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	ARGUMENTSINFO *info = GetArgumentsInfo(this_obj);

	JS_ToInt32(ctx, &info->index, val);
	return JS_UNDEFINED;
}

JSValue ArgumentsatEnd(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ARGUMENTSINFO *info = GetArgumentsInfo(this_obj);

	return JS_NewBool(ctx,
			(info->index >= info->pxc->paramcount - 1));
}

JSValue ArgumentsmoveNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	GetArgumentsInfo(this_obj)->index++;
	return JS_UNDEFINED;
}

JSValue ArgumentsReset(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	GetArgumentsInfo(this_obj)->index = 0;
	return JS_UNDEFINED;
}

JSValue GetArgumentsIteratorDone(JSContext *ctx, JSValueConst this_obj)
{
	ARGUMENTSINFO *info = GetArgumentsInfo(this_obj);

	return JS_NewBool(ctx,
			(info->index >= info->pxc->paramcount - 1));
}

JSValue ArgumentsIteratorNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	GetArgumentsInfo(this_obj)->index++;
	return JS_DupValue(ctx, this_obj);
}

JSValue ArgumentsIterator(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ARGUMENTSINFO *info = GetArgumentsInfo(this_obj);
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Arguments);
	ARGUMENTSINFO *newinfo = malloc(sizeof(ARGUMENTSINFO));
	newinfo->pxc = info->pxc;
	newinfo->index = -1;
	JS_SetOpaque(newobj, newinfo);
	return newobj;
}

void Argumentsfinalizer(JSRuntime *rt, JSValue val)
{
	free(JS_GetOpaque(val, ClassID_Arguments));
}

static const JSCFunctionListEntry ArgumentsFunctionList[] = {
	JS_CFUNC_DEF("Item", 0, ArgumentsItem),
	JS_CFUNC_DEF("item", 0, ArgumentsItem),
	JS_CFUNC_DEF("toString", 0, ArgumentsToString),
	JS_CGETSET_DEF("Count", GetArgumentsCount, NULL),
	JS_CGETSET_DEF("length", GetArgumentsCount, NULL),
	JS_CGETSET_DEF("Index", GetArgumentsIndex, SetArgumentsIndex),
	JS_CGETSET_DEF("index", GetArgumentsIndex, SetArgumentsIndex),
	JS_CGETSET_DEF("value", GetArgumentsValue, NULL),
	JS_CFUNC_DEF("atEnd", 0, ArgumentsatEnd),
	JS_CFUNC_DEF("moveNext", 0, ArgumentsmoveNext),
	JS_CFUNC_DEF("Reset", 0, ArgumentsReset),
#if !MODIFY_QJS
	JS_CFUNC_DEF("Symbol.iterator", 0, ArgumentsIterator),
#endif
	JS_CFUNC_DEF("[Symbol.iterator]", 0, ArgumentsIterator),
	JS_CFUNC_DEF("next", 0, ArgumentsIteratorNext),
	JS_CGETSET_DEF("done", GetArgumentsIteratorDone, NULL),
};

JSClassDef ArgumentsClassDef = {
	"_ArgumentS_", // class_name
	Argumentsfinalizer, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};

static JSValue PPxGetArguments(JSContext *ctx, JSValueConst this_obj)
{
	JSRuntime *rt = JS_GetRuntime(ctx);

	if ( ClassID_Arguments == 0 ) JS_NewClassID(&ClassID_Arguments);
	if ( !JS_IsRegisteredClass(rt, ClassID_Arguments) ){ // クラスを登録
		JS_NewClass(rt, ClassID_Arguments, &ArgumentsClassDef);
		JSValue ProtoClass = JS_NewObject(ctx);
		JS_SetPropertyFunctionList(ctx, ProtoClass, ArgumentsFunctionList, _countof(ArgumentsFunctionList));
		JS_SetClassProto(ctx, ClassID_Arguments, ProtoClass);
	}
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_Arguments);
	ARGUMENTSINFO *info = malloc(sizeof(ARGUMENTSINFO));
	info->pxc = GetCtxPxc(ctx);
	info->index = 0;
	JS_SetOpaque(newobj, info);
	return newobj;
}

//=========================================================== PPx.Enumerator
static JSValue EnumeratorItem(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JSValue jsdata, jsvar, jsresult;
	int index, length;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "index");
	JS_ToInt32(ctx, &index, jsvar);
	JS_FreeValue(ctx, jsvar);

	jsdata = JS_GetPropertyStr(ctx, this_obj, "DATA");
	jsvar = JS_GetPropertyStr(ctx, jsdata, "length");
	JS_ToInt32(ctx, &length, jsvar);
	JS_FreeValue(ctx, jsvar);

	if ( index < length ){
		jsresult = JS_GetPropertyUint32(ctx, jsdata, index);
		JS_FreeValue(ctx, jsdata);
		return jsresult;
	}else{
		JS_FreeValue(ctx, jsdata);
		return JS_NULL;
	}
}

static JSValue EnumeratorAtEnd(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JSValue jsdata, jsvar;
	int index, length;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "index");
	JS_ToInt32(ctx, &index, jsvar);
	JS_FreeValue(ctx, jsvar);

	jsdata = JS_GetPropertyStr(ctx, this_obj, "DATA");
	jsvar = JS_GetPropertyStr(ctx, jsdata, "length");
	if ( !JS_IsNumber(jsvar) ){
		JS_FreeValue(ctx, jsvar);
		jsvar = JS_GetPropertyStr(ctx, jsdata, "Count");
	}
	JS_FreeValue(ctx, jsdata);
	JS_ToInt32(ctx, &length, jsvar);
	JS_FreeValue(ctx, jsvar);

	return JS_NewBool(ctx, (index >= length) );
}

static JSValue EnumeratorMoveNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JSValue jsvar;
	int index;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "index");
	JS_ToInt32(ctx, &index, jsvar);
	JS_FreeValue(ctx, jsvar);
	JS_SetPropertyStr(ctx, this_obj, "index", JS_NewInt32(ctx, index + 1));
	return JS_UNDEFINED;
}

static JSValue EnumeratorMoveFirst(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JS_SetPropertyStr(ctx, this_obj, "index", JS_NewInt32(ctx, 0));
	return JS_UNDEFINED;
}

static const JSCFunctionListEntry EnumeratorFunctionList[] = {
	JS_CFUNC_DEF("item", 1, EnumeratorItem),
	JS_CFUNC_DEF("Item", 1, EnumeratorItem),
	JS_CFUNC_DEF("atEnd", 1, EnumeratorAtEnd),
	JS_CFUNC_DEF("moveNext", 1, EnumeratorMoveNext),
	JS_CFUNC_DEF("moveFirst", 1, EnumeratorMoveFirst),
};

static JSValue IteEnumeratorItem(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return JS_GetPropertyStr(ctx, this_obj, "value");
}

static JSValue IteEnumeratorAtEnd(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return JS_GetPropertyStr(ctx, this_obj, "done");
}

static JSValue IteEnumeratorMoveNext(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JSValue iterator = JS_GetPropertyStr(ctx, this_obj, "iterator");
	JSValue next = JS_GetPropertyStr(ctx, iterator, "next");
	JSValue result = JS_Call(ctx, next, iterator, 0, NULL);
	JS_FreeValue(ctx, next);
	JS_FreeValue(ctx, iterator);

	JSValue var = JS_GetPropertyStr(ctx, result, "done");
	JS_SetPropertyStr(ctx, this_obj, "done", var);

	var = JS_GetPropertyStr(ctx, result, "value");
	JS_SetPropertyStr(ctx, this_obj, "value", var);
	JS_FreeValue(ctx, result);
	return JS_UNDEFINED;
}
/*
static JSValue IteEnumeratorMoveFirst(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JS_SetPropertyStr(ctx, this_obj, "index", JS_NewInt32(ctx, 0));
	return JS_UNDEFINED;
}
*/

static const JSCFunctionListEntry IteEnumeratorFunctionList[] = {
	JS_CFUNC_DEF("item", 1, IteEnumeratorItem),
	JS_CFUNC_DEF("Item", 1, IteEnumeratorItem),
	JS_CFUNC_DEF("atEnd", 1, IteEnumeratorAtEnd),
	JS_CFUNC_DEF("moveNext", 1, IteEnumeratorMoveNext),
//	JS_CFUNC_DEF("moveFirst", 1, IteEnumeratorMoveFirst),
};

static JSValue PPxEnumerator(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JSValue newobj = JS_NewObject(ctx);

#if MODIFY_QJS // ここを有効にするには、quickjs.c/.h の JS_GetIterator の export が必要
	JSValue iterator = JS_GetIterator(ctx, argv[0], FALSE);
	if ( JS_IsObject(iterator) ){
		JS_SetPropertyFunctionList(ctx, newobj, IteEnumeratorFunctionList, _countof(IteEnumeratorFunctionList));
		JS_SetPropertyStr(ctx, newobj, "iterator", iterator);
		IteEnumeratorMoveNext(ctx, newobj, 0, NULL);
	}
#else
	// Object.getOwnPropertySymbols(obj); 経由で取得を検討
	JSValue iterator = JS_GetPropertyStr(ctx, argv[0], "Symbol.iterator");
	if ( JS_IsObject(iterator) ){
		JSValue result = JS_Call(ctx, iterator, argv[0], 0, NULL);
		JS_FreeValue(ctx, iterator);
		JS_SetPropertyFunctionList(ctx, newobj, IteEnumeratorFunctionList, _countof(IteEnumeratorFunctionList));
		JS_SetPropertyStr(ctx, newobj, "iterator", result);
		IteEnumeratorMoveNext(ctx, newobj, 0, NULL);
	}
#endif
	else{
		JS_SetPropertyStr(ctx, newobj, "DATA", JS_DupValue(ctx, argv[0]));
		JS_SetPropertyStr(ctx, newobj, "index", JS_NewInt32(ctx, 0));
		JS_SetPropertyFunctionList(ctx, newobj, EnumeratorFunctionList, _countof(EnumeratorFunctionList));
	}
	return newobj;
}

//=========================================================== PPx objects

void MakeArgsText(JSContext *ctx, ThSTRUCT *text, int argc, JSValueConst *argv)
{
	for ( int i = 0; i < argc; ++i){
		const char *str;

		if ( i != 0 ) THCatStringA(text, " ");
		str = JS_ToCString(ctx, argv[i]);
		if ( str != NULL ) {
			THCatStringA(text, str);
			JS_FreeCString(ctx, str);
		}
	}
}

JSValue PPxEcho(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	ThSTRUCT text = ThSTRUCT_InitData;

	MakeArgsText(ctx, &text, argc, argv);
	PopupMessage(GetCtxPpxa(ctx), (WCHAR *)text.bottom);
	THFree(&text);
	return JS_UNDEFINED;
}

JSValue PPxSetPopLineMessage(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	ThSTRUCT text = ThSTRUCT_InitData;

	MakeArgsText(ctx, &text, argc, argv);
	ppxa->Function(ppxa, PPXCMDID_SETPOPLINE, (WCHAR *)text.bottom);
	THFree(&text);
	return JS_UNDEFINED;
}

JSValue PPxReport(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	ThSTRUCT text = ThSTRUCT_InitData;

	MakeArgsText(ctx, &text, argc, argv);
	ppxa->Function(ppxa, PPXCMDID_REPORTTEXT, (WCHAR *)text.bottom);
	THFree(&text);
	return JS_UNDEFINED;
}

JSValue PPxLog(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	ThSTRUCT text = ThSTRUCT_InitData;

	MakeArgsText(ctx, &text, argc, argv);
	ppxa->Function(ppxa, PPXCMDID_DEBUGLOG, (WCHAR *)text.bottom);
	THFree(&text);
	return JS_UNDEFINED;
}

static JSValue PPxExecute(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	const char *paramA;
	WCHAR *paramW, parambufW[CMDLINESIZE];
	ERRORCODE resultcode;
	int len;

	if ( argc <= 0 ) return JS_UNDEFINED;
	paramA = JS_ToCString(ctx, argv[0]);
	if ( paramA == NULL ) return JS_UNDEFINED;

	len = MultiByteToWideChar(CP_UTF8, 0, paramA, -1, parambufW, CMDLINESIZE);
	if ( len > 0 ){
		resultcode = ppxa->Function(ppxa, PPXCMDID_EXECUTE, parambufW);
	}else{
		len = MultiByteToWideChar(CP_UTF8, 0, paramA, -1, NULL, 0);
		paramW = malloc(len * sizeof(WCHAR));
		if ( paramW != NULL ){
			MultiByteToWideChar(CP_UTF8, 0, paramA, -1, paramW, len);
			resultcode = ppxa->Function(ppxa, PPXCMDID_EXECUTE, paramW);
			free(paramW);
		}else{
			resultcode = ERROR_NOT_ENOUGH_MEMORY;
		}
	}
	if ( resultcode <= 1 ) resultcode ^= 1;
	JS_FreeCString(ctx, paramA);
	return JS_NewInt32(ctx, resultcode);
}

static JSValue PPxExtract(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	InstanceValueStruct *info = GetCtxInfo(ctx);
	const char *paramA;
	WCHAR *paramW, parambufW[CMDLINESIZE];
	BSTR result;
	int len;

	if ( argc == 0 ){
		return JS_NewInt32(ctx, info->extract_result);
	}

	paramA = JS_ToCString(ctx, argv[0]);
	if ( paramA == NULL ) return JS_NULL;

	len = MultiByteToWideChar(CP_UTF8, 0, paramA, -1, parambufW, CMDLINESIZE);
	if ( len > 0 ){
		result = (BSTR)info->ppxa->Function(info->ppxa, PPXCMDID_LONG_EXTRACT_E, parambufW);
	}else{
		len = MultiByteToWideChar(CP_UTF8, 0, paramA, -1, NULL, 0);
		paramW = malloc(len * sizeof(WCHAR));
		if ( paramW != NULL ){
			MultiByteToWideChar(CP_UTF8, 0, paramA, -1, paramW, len);

			result = (BSTR)info->ppxa->Function(info->ppxa, PPXCMDID_LONG_EXTRACT_E, paramW);
			free(paramW);
		}else{
			result = (BSTR)ERROR_NOT_ENOUGH_MEMORY;
		}
	}
	JS_FreeCString(ctx, paramA);

	if ( ((DWORD_PTR)result & ~(DWORD_PTR)0xffff) == 0 ){
		info->extract_result = (ERRORCODE)(DWORD_PTR)result;
		return JS_NewString(ctx, "");
	}else{
		JSValue jsval = JS_NewStringW(ctx, result);
		SysFreeString(result);
		info->extract_result = NO_ERROR;
		return jsval;
	}
}

JSValue PPxSetResult(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	InstanceValueStruct *info = GetCtxInfo(ctx);

	if ( info->pxc->resultstring != NULL ){
		switch ( JS_VALUE_GET_TAG(val) ){
			case JS_TAG_BOOL:
				wcscpy(info->pxc->resultstring, JS_VALUE_GET_BOOL(val) ? L"-1" : L"0");
				break;
			case JS_TAG_NULL:
			case JS_TAG_UNDEFINED:
			case JS_TAG_UNINITIALIZED:
				info->pxc->resultstring[0] = '\0';
				break;
			default:
				const char *str = JS_ToCString(ctx, val);
				if ( str != NULL ){
					int len = MultiByteToWideChar(CP_UTF8, 0, str, -1, info->pxc->resultstring, CMDLINESIZE);
					if ( len == 0 ){
						len = MultiByteToWideChar(CP_UTF8, 0, str, -1, NULL, 0) + 4;
						WCHAR *script = malloc(len * sizeof(WCHAR));
						if ( script != NULL ){
							MultiByteToWideChar(CP_UTF8, 0, str, -1, script, len);
							BSTR LongResult = SysAllocString(script);
							free(script);
							info->ppxa->Function(info->ppxa, PPXCMDID_LONG_RESULT, LongResult);
							info->pxc->resultstring[0] = '\0';
						}
					}
					JS_FreeCString(ctx, str);
				}else{
					info->pxc->resultstring[0] = '\0';
				}
		}
	}
	return JS_UNDEFINED;
}

static JSValue PPxGetResult(JSContext *ctx, JSValueConst this_obj)
{
	return JS_NewStringW(ctx, GetCtxPxc(ctx)->resultstring);
}

static JSValue PPxArgument(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	InstanceValueStruct *info = GetCtxInfo(ctx);
	int index, imax;

	if ( (argc > 0) && !JS_ToInt32(ctx, &index, argv[0]) ){
		if ( index >= 0 ){
			imax = info->pxc->paramcount - 1;
			if ( index < imax ){
				const WCHAR *argptr;

				argptr = info->pxc->param;
				for ( ;;){
					argptr += wcslen(argptr) + 1;
					if ( index-- <= 0 ) break;
				}
				return JS_NewStringW(ctx, argptr);
			}
		}else{
			BSTR param;
			JSValue jsraw;

			param = (BSTR)info->ppxa->Function(info->ppxa, PPXCMDID_GETRAWPARAM, NULL);
			if ( param != NULL ){
				const WCHAR *ptr;
				// 1つ目(ファイル・スクリプト)をスキップ
				ptr = param;
				while ( (*ptr == ' ') || (*ptr == '\t') ) ptr++;
				if ( *ptr == '\"' ){
					ptr++;
					for (;;){
						if ( *ptr == '\0' ) break;
						if ( *ptr++ == '\"' ) break;
					}
					while ( (*ptr == ' ') || (*ptr == '\t') ) ptr++;
				}else {
					for (;;){
						if ( (*ptr == '\0') || (*ptr == ',') ) break;
						ptr++;
					}
				}
				if ( *ptr == ',' ) ptr++;
				jsraw = JS_NewStringW(ctx, ptr);
				SysFreeString(param);
				return jsraw;
			}
		}
	}
	return JS_NULL;
}

static JSValue PPxOption(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
//	InstanceValueStruct *info = GetCtxInfo(ctx);
	JSValue result = JS_NULL;
	const char *name = JS_ToCString(ctx, argv[0]);
	#ifndef _WIN64
	if ( strcmp(name, "Date") == 0 ){
		result = JS_NewBool(ctx, FALSE);
	}
	#endif
	JS_FreeCString(ctx, name);
	return result;
}

static JSValue PPxQuit(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	InstanceValueStruct *info = GetCtxInfo(ctx);

	info->quit = TRUE;
	if ( argc > 0 ) JS_ToInt32(ctx, &info->returncode, argv[0]);
	JS_Throw(ctx, JS_UNDEFINED);
	return JS_EXCEPTION;
}

static JSValue PPxSleep(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	int index;

	if ( !JS_ToInt32(ctx, &index, argv[0]) ) Sleep(index);
	return JS_UNDEFINED;
}

static JSValue PPxLoadCount(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int count = 0;

	JS_ToInt32(ctx, &count, argv[0]);
	ppxa->Function(ppxa, PPXCMDID_DIRLOADCOUNT, &count);
	return JS_NewInt32(ctx, count);
}

static JSValue PPxGetDirectoryType(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_DIRTYPE, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxGetEntryAllCount(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_DIRTOTAL, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxGetEntryDisplayCount(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_DIRTTOTAL, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxGetEntryMarkCount(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_DIRMARKS, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxGetEntryMarkSize(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	__int64 value;

	ppxa->Function(ppxa, PPXCMDID_DIRMARKSIZE, &value);
	return JS_NewInt64(ctx, value);
}

static JSValue PPxGetEntryDisplayDirectories(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_DIRTTOTALDIR, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxGetEntryDisplayFiles(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_DIRTTOTALFILE, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxGetEntryIndex(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_CSRINDEX, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxSetEntryIndex(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value = 0;

	JS_ToInt32(ctx, &value, val);
	ppxa->Function(ppxa, PPXCMDID_CSRSETINDEX, &value);
	return JS_UNDEFINED;
}

static JSValue PPxGetEntryMark(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_CSRMARK, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxSetEntryMark(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value = 0;

	JS_ToInt32(ctx, &value, val);
	ppxa->Function(ppxa, PPXCMDID_CSRSETMARK, &value);
	return JS_UNDEFINED;
}

static JSValue PPxGetEntryName(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR bufw[CMDLINESIZE];
	PPXCMDENUMSTRUCTW info;

	info.buffer = bufw;
	ppxa->Function(ppxa, 'R', &info);
	return JS_NewStringW(ctx, bufw);
}

static JSValue PPxGetEntryAttributes(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value;

	ppxa->Function(ppxa, PPXCMDID_CSRATTR, &value);
	return JS_NewInt32(ctx, value);
}

static JSValue PPxGetEntrySize(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	__int64 value;

	ppxa->Function(ppxa, PPXCMDID_CSRMSIZE, &value);
	return JS_NewInt64(ctx, value);
}

static JSValue PPxGetEntryState(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	ppxa->Function(ppxa, PPXCMDID_CSRSTATE, &index);
	return JS_NewInt32(ctx, index);
}

static JSValue PPxSetEntryState(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int tmpvalue;
	int value = 0;

	JS_ToInt32(ctx, &value, val);
	ppxa->Function(ppxa, PPXCMDID_CSRSTATE, &tmpvalue);
	value = (tmpvalue & 0xffe0) | value;
	ppxa->Function(ppxa, PPXCMDID_CSRSETSTATE, &value);
	return JS_UNDEFINED;
}

static JSValue PPxGetEntryExtColor(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	ppxa->Function(ppxa, PPXCMDID_CSREXTCOLOR, &index);
	return JS_NewInt32(ctx, index);
}

static JSValue PPxSetEntryExtColor(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int value = 0;

	JS_ToInt32(ctx, &value, val);
	ppxa->Function(ppxa, PPXCMDID_CSRSETEXTCOLOR, &value);
	return JS_UNDEFINED;
}

static JSValue PPxGetEntryHighlight(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	ppxa->Function(ppxa, PPXCMDID_CSRSTATE, &index);
	return JS_NewInt32(ctx, index >> 5);
}

static JSValue PPxSetEntryHighlight(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int tmpvalue;
	int value = 0;

	JS_ToInt32(ctx, &value, val);
	ppxa->Function(ppxa, PPXCMDID_CSRSTATE, &tmpvalue);
	value = (tmpvalue & 0x1f) | (value << 5);
	ppxa->Function(ppxa, PPXCMDID_CSRSETSTATE, &value);
	return JS_UNDEFINED;
}

static JSValue PPxGetEntryComment(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR bufw[CMDLINESIZE];

	ppxa->Function(ppxa, PPXCMDID_CSRCOMMENT, bufw);
	return JS_NewStringW(ctx, bufw);
}

static JSValue PPxSetEntryComment(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR commentbuf[CMDLINESIZE], *comment;

	comment = GetJsLongString(ctx, val, commentbuf);
	ppxa->Function(ppxa, PPXCMDID_CSRSETCOMMENT, (void*)comment);
	FreeJsLongString(comment, commentbuf);
	return JS_UNDEFINED;
}

static JSValue PPxGetEntryFirstMark(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	if ( ppxa->Function(ppxa, PPXCMDID_CSRMARKFIRST, &index) == PPXA_NO_ERROR ){
		return JS_NewInt32(ctx, index);
	}
	JS_ThrowReferenceError(ctx, "entry not support");
	return JS_EXCEPTION;
}

static JSValue PPxGetEntryNextMark(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	if ( ppxa->Function(ppxa, PPXCMDID_CSRMARKNEXT, &index) == PPXA_NO_ERROR ){
		return JS_NewInt32(ctx, index);
	}
	JS_ThrowReferenceError(ctx, "entry not support");
	return JS_EXCEPTION;
}

static JSValue PPxGetEntryPrevMark(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	ppxa->Function(ppxa, PPXCMDID_CSRMARKPREV, &index);
	return JS_NewInt32(ctx, index);
}

static JSValue PPxGetEntryLastMark(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	ppxa->Function(ppxa, PPXCMDID_CSRMARKLAST, &index);
	return JS_NewInt32(ctx, index);
}

static JSValue PPxGetPointType(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int var[3];

	ppxa->Function(ppxa, PPXCMDID_POINTINFO, &var);
	return JS_NewInt32(ctx, var[0]);
}

static JSValue PPxGetPointIndex(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int var[3];

	ppxa->Function(ppxa, PPXCMDID_POINTINFO, &var);
	return JS_NewInt32(ctx, var[1]);
}

static JSValue PPxGetStayMode(JSContext *ctx, JSValueConst this_obj)
{
	InstanceValueStruct *info = GetCtxInfo(ctx);
	return JS_NewInt32(ctx, info->stay.mode);
}

static JSValue PPxSetStayMode(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	InstanceValueStruct *info = GetCtxInfo(ctx);
	int value = 0;

	JS_ToInt32(ctx, &value, val);
	if ( value < 0 ){
		JS_ThrowReferenceError(ctx, "StayMode >= 0");
		return JS_EXCEPTION;
	}

	if ( value >= ScriptStay_Stay ){
		if ( info->stay.threadID == 0 ){
			if ( value == ScriptStay_Stay ){
				value = StayInstanseIDserver++;
				if ( value >= ScriptStay_MaxAutoID ){
					StayInstanseIDserver = ScriptStay_FirstAutoID;
				}
			}
			ChainStayInstance(info);
		}
	}else{
		if ( info->stay.threadID != 0 ){
			DropStayInstance(info);
		}
	}
	info->stay.mode = value;
	return JS_UNDEFINED;
}

static JSValue PPxGetCodeType(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	return JS_NewInt32(ctx, ppxa->Function(ppxa, PPXCMDID_CHARTYPE, NULL));
}

static JSValue PPxGetPPxVersion(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	return JS_NewInt32(ctx, ppxa->Function(ppxa, PPXCMDID_VERSION, NULL));
}

static JSValue PPxGetWindowIDName(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
#if 0
	return JS_NewStringW(ctx, ppxa->RegID);
#else
	WCHAR bufW[REGEXTIDSIZE];

	if ( (ppxa->Function(ppxa, PPXCMDID_GETSUBID, &bufW) == 0) || ((UTCHAR)bufW[0] > 'z') ){
		if ( ppxa->RegID[1] != '_' ){
			return JS_NewStringW(ctx, ppxa->RegID);
		}
		bufW[0] = ppxa->RegID[0];
		strcpyW(bufW + 1, ppxa->RegID + 2);
	}
	return JS_NewStringW(ctx, bufW);
#endif
}

static JSValue PPxSetWindowIDName(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
//	なにもしない
	return JS_UNDEFINED;
}

static JSValue PPxGetWindowDirection(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int val;
	ppxa->Function(ppxa, PPXCMDID_WINDOWDIR, &val);
	return JS_NewInt32(ctx, val);
}

static JSValue PPxGetEntryDisplayTop(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int val;
	ppxa->Function(ppxa, PPXCMDID_CSRDINDEX, &val);
	return JS_NewInt32(ctx, val);
}

static JSValue PPxSetEntryDisplayTop(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	if ( !JS_ToInt32(ctx, &index, val) ){
		ppxa->Function(ppxa, PPXCMDID_CSRSETDINDEX, &index);
	}
	return JS_UNDEFINED;
}

static JSValue PPxGetEntryDisplayX(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int val;
	ppxa->Function(ppxa, PPXCMDID_CSRDISPW, &val);
	return JS_NewInt32(ctx, val);
}

static JSValue PPxGetEntryDisplayY(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int val;
	ppxa->Function(ppxa, PPXCMDID_CSRDISPH, &val);
	return JS_NewInt32(ctx, val);
}

static JSValue PPxGetComboItemCount(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index = 0;
	if ( argc > 0 ) JS_ToInt32(ctx, &index, argv[0]);
	ppxa->Function(ppxa, PPXCMDID_COMBOIDCOUNT, &index);
	return JS_NewInt32(ctx, index);
}

static JSValue PPxGetSlowMode(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int val;
	ppxa->Function(ppxa, PPXCMDID_SLOWMODE, &val);
	return JS_NewInt32(ctx, val);
}

static JSValue PPxSetSlowMode(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	if ( !JS_ToInt32(ctx, &index, val) ){
		ppxa->Function(ppxa, PPXCMDID_SETSLOWMODE, &index);
	}
	return JS_UNDEFINED;
}

static JSValue PPxGetSyncView(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int val;
	ppxa->Function(ppxa, PPXCMDID_SYNCVIEW, &val);
	return JS_NewInt32(ctx, val);
}

static JSValue PPxSetSyncView(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	int index;

	if ( !JS_ToInt32(ctx, &index, val) ){
		ppxa->Function(ppxa, PPXCMDID_SETSYNCVIEW, &index);
	}
	return JS_UNDEFINED;
}

static JSValue PPxGetDriveVolumeLabel(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR bufw[CMDLINESIZE];

	bufw[0] = '\0';
	ppxa->Function(ppxa, PPXCMDID_DRIVELABEL, bufw);
	return JS_NewStringW(ctx, bufw);
}

static JSValue PPxGetDriveTotalSize(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	__int64 size;

	ppxa->Function(ppxa, PPXCMDID_DRIVETOTALSIZE, &size);
	return JS_NewInt64(ctx, size);
}

static JSValue PPxGetDriveFreeSize(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	__int64 size;

	ppxa->Function(ppxa, PPXCMDID_DRIVEFREE, &size);
	return JS_NewInt64(ctx, size);
}

static JSValue PPxGetScriptName(JSContext *ctx, JSValueConst this_obj)
{
	return JS_NewStringW(ctx, GetCtxPxc(ctx)->param);
}

JSValue PPxGetFileInformation(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	PPXUPTR_FILEINFORMATION tmp;

	tmp.type = 0;
	GetJsShortString(ctx, argv[0], tmp.filename);
	if ( argc > 1 ) JS_ToInt32(ctx, &tmp.type, argv[1]);
	ppxa->Function(ppxa, PPXCMDID_GETFILEINFO, &tmp);
	return JS_NewStringW(ctx, tmp.filename);
}

JSValue PPxEntryInsert(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	DWORD_PTR dptrs[2];
	WCHAR entryname[CMDLINESIZE];

	JS_ToInt64(ctx, (int64_t *)&dptrs[0], argv[0]);
	GetJsShortString(ctx, argv[1], entryname);

	dptrs[1] = (DWORD_PTR)entryname;
	ppxa->Function(ppxa, PPXCMDID_ENTRYINSERTMSG, dptrs);
	return JS_UNDEFINED;
}

JSValue PPxgetValue(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR *uptr[2];
	WCHAR key[CMDLINESIZE];
	WCHAR buf[CMDLINESIZE];

	GetJsShortString(ctx, argv[0], key);
	uptr[0] = key;
	uptr[1] = buf;
	if ( 0 == ppxa->Function(ppxa, PPXCMDID_GETPROCVARIABLEDATA, uptr) ){
		buf[0] = '\0';
	}
	return JS_NewStringW(ctx, buf);
}

JSValue PPxsetValue(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR *uptr[2];
	WCHAR key[CMDLINESIZE];
	WCHAR valuebuf[CMDLINESIZE], *value;

	GetJsShortString(ctx, argv[0], key);
	value = GetJsLongString(ctx, argv[1], valuebuf);
	uptr[0] = key;
	uptr[1] = value;
	ppxa->Function(ppxa, PPXCMDID_SETPROCVARIABLEDATA, uptr);
	FreeJsLongString(value, valuebuf);
	return JS_UNDEFINED;
}

JSValue PPxgetIValue(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR *uptr[2];
	WCHAR key[CMDLINESIZE];
	WCHAR buf[CMDLINESIZE];

	GetJsShortString(ctx, argv[0], key);
	uptr[0] = key;
	uptr[1] = buf;
	if ( 0 == ppxa->Function(ppxa, PPXCMDID_GETWNDVARIABLEDATA, uptr) ){
		buf[0] = '\0';
	}
	return JS_NewStringW(ctx, buf);
}

JSValue PPxsetIValue(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR *uptr[2];
	WCHAR key[CMDLINESIZE];
	WCHAR valuebuf[CMDLINESIZE], *value;

	GetJsShortString(ctx, argv[0], key);
	value = GetJsLongString(ctx, argv[1], valuebuf);
	uptr[0] = key;
	uptr[1] = value;
	ppxa->Function(ppxa, PPXCMDID_SETWNDVARIABLEDATA, uptr);
	FreeJsLongString(value, valuebuf);
	return JS_UNDEFINED;
}

JSValue PPxInclude(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR filename[CMDLINESIZE];
	char filenameA[CMDLINESIZE];
	char *script;
	JSValue result;
	DWORD script_len;
	char *TextImage;

	GetJsShortString(ctx, argv[0], filename);
	UnicodeToAnsi(filename,filenameA,1000);
	script = LoadScript(filename, &TextImage, &script_len);
	result = JS_Eval(ctx, script, script_len, filenameA, JS_EVAL_TYPE_GLOBAL);
	free(TextImage);
	return result;
}

JSValue PPxgetReentryCount(JSContext *ctx, JSValueConst this_obj)
{
	InstanceValueStruct *info = GetCtxInfo(ctx);
	return JS_NewInt32(ctx, info->stay.entry);
}

BOOL TryOpenClipboard(HWND hWnd)
{
	int trycount = 6;

	for (;;){
		if ( IsTrue(OpenClipboard(hWnd)) ) return TRUE;
		if ( --trycount == 0 ) return FALSE;
		Sleep(20);
	}
}

static JSValue PPxGetClipboard(JSContext *ctx, JSValueConst this_obj)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	JSValue jsclipdata = JS_UNDEFINED;

	if ( TryOpenClipboard(ppxa->hWnd) != FALSE ){
		HANDLE clipdata;
		WCHAR *src;

		clipdata = GetClipboardData(CF_UNICODETEXT);
		if ( clipdata != NULL ){
			src = (WCHAR *)GlobalLock(clipdata);
			jsclipdata = JS_NewStringW(ctx, src);
			GlobalUnlock(clipdata);
		}
		CloseClipboard();
	}
	return jsclipdata;
}

static JSValue PPxSetClipboard(JSContext *ctx, JSValueConst this_obj, JSValueConst val)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR bufW[CMDLINESIZE], *str;
	HGLOBAL hm;
	size_t size;

	str = GetJsLongString(ctx, val, bufW);
	size = wcslen(str) * sizeof(WCHAR) + sizeof(WCHAR);
	hm = GlobalAlloc(GMEM_MOVEABLE, size);
	if ( hm != NULL ){
		memcpy(GlobalLock(hm), str, size);
		GlobalUnlock(hm);
		TryOpenClipboard(ppxa->hWnd);
		EmptyClipboard();
		SetClipboardData(CF_UNICODETEXT, hm);
		CloseClipboard();
	}
	FreeJsLongString(str, bufW);
	return JS_UNDEFINED;
}

const JSCFunctionListEntry PPxFunctionList[] = {
	JS_CGETSET_DEF("Arguments", PPxGetArguments, NULL),
	JS_CFUNC_DEF("Argument", 1, PPxArgument),
	JS_CFUNC_DEF("Echo", 0, PPxEcho),
	JS_CFUNC_DEF("SetPopLineMessage", 0, PPxSetPopLineMessage),
	JS_CFUNC_DEF("linemessage", 0, PPxSetPopLineMessage),
	JS_CFUNC_DEF("report", 0, PPxReport),
	JS_CFUNC_DEF("log", 0, PPxLog),
	JS_CFUNC_DEF("CreateObject", 1, PPxCreateObject),
	JS_CFUNC_DEF("GetObject", 1, PPxGetObject),
	JS_CFUNC_MAGIC_DEF("ConnectObject", 2, PPxConnectObject, 0),
	JS_CFUNC_MAGIC_DEF("DisconnectObject", 1, PPxConnectObject, 1),
	JS_CFUNC_DEF("Enumerator", 1, PPxEnumerator),
	JS_CFUNC_DEF("Execute", 1, PPxExecute),
	JS_CFUNC_DEF("Extract", 0, PPxExtract),
	JS_CFUNC_DEF("option", 0, PPxOption),
	JS_CFUNC_DEF("Quit", 0, PPxQuit),
	JS_CFUNC_DEF("quit", 0, PPxQuit),
	JS_CFUNC_DEF("Sleep", 0, PPxSleep),
	JS_CFUNC_DEF("GetFileInformation", 1, PPxGetFileInformation),
	JS_CGETSET_DEF("Clipboard", PPxGetClipboard, PPxSetClipboard),
	JS_CFUNC_DEF("EntryInsert", 1, PPxEntryInsert),
	JS_CFUNC_DEF("setValue", 2, PPxsetValue),
	JS_CFUNC_DEF("getValue", 1, PPxgetValue),
	JS_CFUNC_DEF("setProcessValue", 2, PPxsetValue),
	JS_CFUNC_DEF("getProcessValue", 1, PPxgetValue),
	JS_CFUNC_DEF("setIValue", 2, PPxsetIValue),
	JS_CFUNC_DEF("getIValue", 1, PPxgetIValue),
	JS_CFUNC_DEF("Include", 0, PPxInclude),
	JS_CGETSET_DEF("ReentryCount", PPxgetReentryCount, NULL),

	JS_CGETSET_DEF("Pane", PPxGetPane, NULL),
	JS_CGETSET_DEF("Entry", PPxGetEntry, NULL),

	JS_CFUNC_DEF("LoadCount", 0, PPxLoadCount),
	JS_CGETSET_DEF("DirectoryType", PPxGetDirectoryType, NULL),
	JS_CGETSET_DEF("EntryAllCount", PPxGetEntryAllCount, NULL),
	JS_CGETSET_DEF("EntryDisplayCount", PPxGetEntryDisplayCount, NULL),
	JS_CGETSET_DEF("EntryMarkCount", PPxGetEntryMarkCount, NULL),
	JS_CGETSET_DEF("EntryMarkSize", PPxGetEntryMarkSize, NULL),
	JS_CGETSET_DEF("EntryDisplayDirectories", PPxGetEntryDisplayDirectories, NULL),
	JS_CGETSET_DEF("EntryDisplayFiles", PPxGetEntryDisplayFiles, NULL),

	JS_CGETSET_DEF("EntryIndex", PPxGetEntryIndex, PPxSetEntryIndex),
	JS_CGETSET_DEF("EntryMark", PPxGetEntryMark, PPxSetEntryMark),
	JS_CGETSET_DEF("EntryName", PPxGetEntryName, NULL),
	JS_CGETSET_DEF("EntryAttributes", PPxGetEntryAttributes, NULL),
	JS_CGETSET_DEF("EntrySize", PPxGetEntrySize, NULL),
	JS_CGETSET_DEF("EntryState", PPxGetEntryState, PPxSetEntryState),
	JS_CGETSET_DEF("EntryExtColor", PPxGetEntryExtColor, PPxSetEntryExtColor),
	JS_CGETSET_DEF("EntryHighlight", PPxGetEntryHighlight, PPxSetEntryHighlight),
	JS_CGETSET_DEF("EntryComment", PPxGetEntryComment, PPxSetEntryComment),
	JS_CGETSET_DEF("EntryFirstMark", PPxGetEntryFirstMark, NULL),
	JS_CGETSET_DEF("EntryNextMark", PPxGetEntryNextMark, NULL),
	JS_CGETSET_DEF("EntryPrevMark", PPxGetEntryPrevMark, NULL),
	JS_CGETSET_DEF("EntryLastMark", PPxGetEntryLastMark, NULL),
	JS_CGETSET_DEF("PointType", PPxGetPointType, NULL),
	JS_CGETSET_DEF("PointIndex", PPxGetPointIndex, NULL),
	JS_CGETSET_DEF("StayMode", PPxGetStayMode, PPxSetStayMode),
	JS_CGETSET_DEF("ScriptName", PPxGetScriptName, NULL),
	JS_CGETSET_DEF("ScriptFullName", PPxGetScriptName, NULL),
	JS_CGETSET_DEF("result", PPxGetResult, PPxSetResult),
	JS_CGETSET_DEF("CodeType", PPxGetCodeType, NULL),
	JS_CGETSET_DEF("PPxVersion", PPxGetPPxVersion, NULL),
	JS_CGETSET_DEF("WindowIDName", PPxGetWindowIDName, PPxSetWindowIDName),
	JS_CGETSET_DEF("WindowDirection", PPxGetWindowDirection, NULL),
	JS_CGETSET_DEF("EntryDisplayTop", PPxGetEntryDisplayTop, PPxSetEntryDisplayTop),
	JS_CGETSET_DEF("EntryDisplayX", PPxGetEntryDisplayX, NULL),
	JS_CGETSET_DEF("EntryDisplayY", PPxGetEntryDisplayY, NULL),
	JS_CFUNC_DEF("GetComboItemCount", 0, PPxGetComboItemCount),
	JS_CGETSET_DEF("SlowMode", PPxGetSlowMode, PPxSetSlowMode),
	JS_CGETSET_DEF("SyncView", PPxGetSyncView, PPxSetSyncView),
	JS_CGETSET_DEF("DriveVolumeLabel", PPxGetDriveVolumeLabel, NULL),
	JS_CGETSET_DEF("DriveTotalSize", PPxGetDriveTotalSize, NULL),
	JS_CGETSET_DEF("DriveFreeSize", PPxGetDriveFreeSize, NULL),
	JS_PROP_INT32_DEF("ModuleVersion", SCRIPTMODULEVER, 0),
	JS_PROP_STRING_DEF("ScriptEngineName", "QuickJS", 0),
	JS_PROP_STRING_DEF("ScriptEngineVersion", QUICKJSVERSION, 0),

	JS_PROP_INT32_DEF("emuole", 0, JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
};

JSClassDef PPxClassDef = {
	"_PpX_", // class_name
	NULL, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};

JSModuleDef *ModuleLoader(JSContext *ctx, const char *filenameA, void *opaque)
{
	JSModuleDef *module;
	JSValue funcbin;

	char *script;
	DWORD sizeL;		// ファイルの大きさ
	char *TextImage;
	WCHAR filenameW[CMDLINESIZE];

	MultiByteToWideChar(CP_UTF8, 0, filenameA, -1, filenameW, CMDLINESIZE);
	script = LoadScript(filenameW, &TextImage, &sizeL);
	if ( script == NULL ){
		JS_ThrowReferenceError(ctx, "could not load module filename '%s'", filenameA);
		return NULL;
	}
	funcbin = JS_Eval(ctx, script, sizeL, filenameA,
			JS_EVAL_TYPE_MODULE | JS_EVAL_FLAG_COMPILE_ONLY);
	free(TextImage);

	if ( JS_IsException(funcbin) ) return NULL;
	SetEvalInfo(ctx, funcbin, FALSE);
	module = JS_VALUE_GET_PTR(funcbin);
	JS_FreeValue(ctx, funcbin);
	return module;
}

JSContext *RunScriptNewContext(JSRuntime *rt, InstanceValueStruct *info)
{
	JSContext *ctx;
	JSValue global;

	ctx = JS_NewContext(rt);
	JS_SetContextOpaque(ctx, info);
	global = JS_GetGlobalObject(ctx);

	if ( ClassID_PPx == 0 ) JS_NewClassID(&ClassID_PPx);
	if ( !JS_IsRegisteredClass(rt, ClassID_PPx) ){ // クラスを登録
		JS_NewClass(rt, ClassID_PPx, &PPxClassDef);
		JSValue ProtoClass = JS_NewObject(ctx);
		JS_SetPropertyFunctionList(ctx, ProtoClass, PPxFunctionList, _countof(PPxFunctionList));
		JS_SetClassProto(ctx, ClassID_PPx, ProtoClass);
	}
	JSValue newobj = JS_NewObjectClass(ctx, ClassID_PPx);
	JS_SetPropertyStr(ctx, global, "PPx", newobj);

	JS_SetModuleLoaderFunc(rt, NULL, ModuleLoader, NULL);

	JS_FreeValue(ctx, global);
	return ctx;
}
