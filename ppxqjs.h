/*-----------------------------------------------------------------------------
	Paper Plane xUI QuickJS Script Module							(c)TORO
-----------------------------------------------------------------------------*/
#define TOSTRMACRO(item)	#item

#define QUICKJSVERSION "2024-01-13"
#define SCRIPTMODULEVER 0  // Release number
#define SCRIPTMODULEVERSTR	UNICODESTR(TOSTRMACRO(0))

#define Message8(text) {WCHAR bufW[2000]; MultiByteToWideChar(CP_UTF8, 0, text, -1, bufW, 2000); MessageW(bufW); }

#ifdef __cplusplus
extern "C" {
#endif

typedef struct tagFREECHAIN {
	struct tagFREECHAIN *next;
	void *value;
} FREECHAIN;
#define Chain_FreeChainObject(rootchain, chain, object) { (chain).value = (object); (chain).next = (rootchain).next; (rootchain).next = &(chain);}

// Instance value(JSRuntime 内に配置)
typedef struct {
	PPXAPPINFOW *ppxa;
	PPXMCOMMANDSTRUCT *pxc;
	FREECHAIN EventChain;
	JSRuntime *rt;
	JSContext *ctx;
	char *script_image; // root のソース。エラー報告時に参照している

	struct { // 常駐時に必要な要素
		PPXAPPINFOW *ppxa; // 非アクティブ時に使用する ppxa 要 HeapFree
		HWND hWnd;
		WCHAR *source;
		FREECHAIN chain;
		#define ScriptStay_None		0
		#define ScriptStay_Cache	1
		#define ScriptStay_Stay		2
		// ID 0-9999 setting
		// ID 2000-999999999 user
		#define ScriptStay_FirstUserID 2000
		#define ScriptStay_FirstAutoID 0x40000000
		#define ScriptStay_MaxAutoID 0x7fffff00
		int mode;

		DWORD threadID;
		int entry; // 使用中カウンタ
	} stay;

	ERRORCODE extract_result;
	int returncode;
	BOOL quit; // PPx.Quit 中のため、エラー表示しない
} InstanceValueStruct;

#define GetCtxInfo(ctx) ((InstanceValueStruct *)JS_GetContextOpaque(ctx))
#define GetCtxPxc(ctx) (GetCtxInfo(ctx)->pxc)
#define GetCtxPpxa(ctx) (GetCtxInfo(ctx)->ppxa)

extern HANDLE ProcHeap; // GetProcessHeap() の値

//===========================================================
extern void PopupMessage(PPXAPPINFOW *ppxa, const WCHAR *message);
extern void PopupMessagef(PPXAPPINFOW *ppxa, const WCHAR *format, ...);

extern void THInit(ThSTRUCT *TH);
extern BOOL THFree(ThSTRUCT *TH);
extern BOOL THCatStringW(ThSTRUCT *TH, const WCHAR *data);
extern BOOL THCatStringA(ThSTRUCT *TH, const char *data);

//=========================================================== 型変換
extern JSValue JS_NewStringW(JSContext *ctx, const WCHAR *str);
#define GetJsShortString(ctx, js_var, varW) {const char *varA; varA = JS_ToCString(ctx, js_var); if ( varA != NULL ){ MultiByteToWideChar(CP_UTF8, 0, varA, -1, (varW), CMDLINESIZE); JS_FreeCString(ctx, varA); } else { (varW)[0] = '\0'; }}
extern WCHAR *GetJsLongString(JSContext *ctx, JSValueConst js_var, WCHAR *bufW);
#define FreeJsLongString(valW, bufW) {if ( (valW) != (bufW) ){ free(valW); }}

//===========================================================
extern void ChainStayInstance(InstanceValueStruct *info);
extern void DropStayInstance(InstanceValueStruct *info);
extern int SetEvalInfo(JSContext *ctx, JSValueConst func_val, JS_BOOL is_main);
extern char *LoadScript(const WCHAR *filename, char **TextImage, DWORD *sizeL);

extern void ReleaseAllEvent(FREECHAIN *chain);

// qjs_PPxObjects
extern JSContext *RunScriptNewContext(JSRuntime *rt, InstanceValueStruct *info);
extern JSValue PPxSetResult(JSContext *ctx, JSValueConst this_obj, JSValueConst val);

// qjs_CreateObject
extern JSValue PPxCreateObject(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv);
extern JSValue PPxGetObject(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv);
extern JSValue PPxConnectObject(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv, int magic);

// qjs_OLEobject
extern JSValue New_ADODB_Stream(JSContext *ctx);
extern JSValue New_FileSystemObject(JSContext *ctx);

#ifndef RELESE
extern int CountDisp, CountDispRelease, CountEvent, CountEventRelease;
#endif

#ifdef __cplusplus
	}
#endif
