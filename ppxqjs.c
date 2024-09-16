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

typedef struct { // ppxa とかの復旧用
	PPXAPPINFOW* ppxa;
	PPXMCOMMANDSTRUCT* pxc;
} OldPPxInfoStruct;

HANDLE ProcHeap; // GetProcessHeap() の値

long PPxVersion = 0; // 常駐機能有効の判断用
FREECHAIN StayChains = {NULL, NULL}; // 常駐用の管理
CRITICAL_SECTION StayLock;

int RunScript(PPXAPPINFOW *ppxa, PPXMCOMMANDSTRUCT *pxc, int file, BOOL module);
int StayInfo(PPXAPPINFOW *ppxa, PPXMCOMMANDSTRUCT *pxc);
void FreeStayInstance(PPXAPPINFOW *ppxa, BOOL terminate);
void CloseInstance(InstanceValueStruct *info);
extern PPXAPPINFOW DummyPPxAppInfo;

BOOL WINAPI DllMain(HINSTANCE hInst, DWORD reason, LPVOID reserved)
{
	if ( reason == DLL_PROCESS_DETACH ){
		if ( PPxVersion != 0 ) DeleteCriticalSection(&StayLock);
	}
	return TRUE;
}

__declspec(dllexport) int PPXAPI ModuleEntry(PPXAPPINFOW *ppxa, DWORD cmdID, PPXMODULEPARAM pxs)
{
	if ( (cmdID == PPXMEVENT_COMMAND) || (cmdID == PPXMEVENT_FUNCTION) ){
/*
		pxs.command->commandhash を使うと、コマンド名判別の高速化ができる。
		もう一方の値は PPx で " *commandhash コマンド名 " を実行すると分かる。
		※コマンド名が表示されているときに Ctrl+C でコピー可能
*/
		// ====================================================================
		if ( ((pxs.command->commandhash == 0x8004b4d1) && !wcscmp(pxs.command->commandname, L"JSQ")) ||
			 ((pxs.command->commandhash == 0x800012d3) && !wcscmp(pxs.command->commandname, L"JS")) ){
			return RunScript(ppxa, pxs.command, 0, FALSE);
		}
		if ( ((pxs.command->commandhash == 0xd3251571) && !wcscmp(pxs.command->commandname, L"SCRIPTQ")) ||
			 ((pxs.command->commandhash == 0xc34c9454) && !wcscmp(pxs.command->commandname, L"SCRIPT")) ){
			return RunScript(ppxa, pxs.command, 1, FALSE);
		}
		if ( (pxs.command->commandhash == 0xd325157d) && !wcscmp(pxs.command->commandname, L"SCRIPTM") ){
			return RunScript(ppxa, pxs.command, 1, TRUE);
		}
		if ( (pxs.command->commandhash == 0xd925fddf) &&
			 (cmdID == PPXMEVENT_FUNCTION) &&
			 !wcscmp(pxs.command->commandname, L"STAYINFO") ){
			return StayInfo(ppxa, pxs.command);
		}
		return PPXMRESULT_SKIP;
	}

	if ( cmdID == PPXMEVENT_CLEANUP ){
		// 1.97 以前は CleanUp 時に CoUninitialize() になってることがあるので無効に
		if ( PPxVersion > 19700 ) FreeStayInstance(&DummyPPxAppInfo, TRUE);
		return PPXMRESULT_DONE;
	}

	if ( cmdID == PPXMEVENT_CLOSETHREAD ){ // 1.97+1 から有効
		FreeStayInstance(ppxa, TRUE);
		return PPXMRESULT_DONE;
	}

	if ( cmdID == PPXMEVENT_DESTROY ){ // 1.98+4 から有効
		FreeStayInstance(ppxa, FALSE);
		return PPXMRESULT_DONE;
	}

	if ( cmdID == PPXM_INFORMATION ){
		if ( pxs.info->infotype == 0 ){
			pxs.info->typeflags = PPMTYPEFLAGS(PPXM_INFORMATION) |
					PPMTYPEFLAGS(PPXMEVENT_CLEANUP) |
					PPMTYPEFLAGS(PPXMEVENT_CLOSETHREAD) |
					PPMTYPEFLAGS(PPXMEVENT_DESTROY) |
					PPMTYPEFLAGS(PPXMEVENT_COMMAND) |
					PPMTYPEFLAGS(PPXMEVENT_FUNCTION);
			wcscpy(pxs.info->copyright, L"PPx QucickJS script Module T" SCRIPTMODULEVERSTR L"  Copyright (c)TORO");
			return PPXMRESULT_DONE;
		}
	}
	return PPXMRESULT_SKIP;
}

// 異常終了したときにリビジョンを表示させるためのダミーAPI
#define SCRVERNAMEAPI(Ver) SCRVERNAMEAPI2(Ver)
#define SCRVERNAMEAPI2(Ver) PPxQjsT##Ver
__declspec(dllexport) BOOL PPXAPI SCRVERNAMEAPI(SCRIPTMODULEVER)(void)
{
	return FALSE;
}

/*-----------------------------------------------------------------------------
	メッセージダイアログを表示（console時は、teminal内で表示される）
-----------------------------------------------------------------------------*/
void PopupMessage(PPXAPPINFOW *ppxa, const WCHAR *message)
{
	if ( (ppxa == NULL) ||
		 (ppxa->Function == NULL) ||
		 (ppxa->Function(ppxa, PPXCMDID_MESSAGE, (void *)message) == 0) ){
		HWND hWnd = (ppxa == NULL) ? NULL : ppxa->hWnd;
		MessageBoxW(hWnd, message, L"PPxQjs.dll", MB_OK);
	}
}

void PopupMessagef(PPXAPPINFOW *ppxa, const WCHAR *format, ...)
{
	WCHAR message[1024 + 16];
	va_list argptr;
	va_start(argptr, format);
	wvsprintfW(message, format, argptr);
	va_end(argptr);
	message[1024] = '\0';
	PopupMessage(ppxa, message);
}

//=========================================================== Stack Heap 関連
void THInit(ThSTRUCT *TH)
{
	TH->bottom = NULL;
	TH->top = 0;
	TH->size = 0;
}

BOOL THFree(ThSTRUCT *TH)
{
	BOOL result = TRUE;
	char *bottom;

	bottom = TH->bottom;
	if ( bottom != NULL ){
		TH->bottom = NULL;
		TH->top = 0;
		TH->size = 0;
		result = HeapFree(ProcHeap, 0, bottom);
	}
	return result;
}

#define ThAllocCheck()								\
	if ( TH->bottom == NULL ){						\
		TH->bottom = HeapAlloc(ProcHeap, 0, ThSTEP);	\
		if ( TH->bottom == NULL ) return FALSE;		\
		TH->size = ThSTEP;							\
	}

#define ThSizecheck(CHECKSIZE)						\
	while ( (TH->top + (CHECKSIZE)) > TH->size ){	\
		char *p;									\
		DWORD nextsize;								\
													\
		nextsize = TH->size + ThNextAllocSizeM(TH->size < (256 * KB) ? TH->size : (TH->size * 2) );\
		p = HeapReAlloc(ProcHeap, 0, TH->bottom, nextsize);\
		if (p == NULL) return FALSE;				\
		TH->bottom = (void *)p;						\
		TH->size = nextsize;						\
	}

BOOL THSize(ThSTRUCT *TH, DWORD size)
{
	ThAllocCheck();
	ThSizecheck(size);
	return TRUE;
}

BOOL THAppend(ThSTRUCT *TH, const void *data, DWORD size)
{
	ThAllocCheck();
	ThSizecheck(size);
	memcpy(ThLast(TH), data, size);
	TH->top += size;
	return TRUE;
}

BOOL THCatStringW(ThSTRUCT *TH, const WCHAR *data)
{
	DWORD size;

	ThAllocCheck();
	size = (strlenW(data) + 1) * sizeof(WCHAR);
	ThSizecheck(size);
	memcpy(ThLast(TH), data, size);
	TH->top += size - sizeof(WCHAR);
	return TRUE;
}

BOOL THCatStringA(ThSTRUCT *TH, const char *data)
{
	DWORD size;

	ThAllocCheck();
	size = MultiByteToWideChar(CP_UTF8, 0, data, -1, NULL, 0);
	ThSizecheck(size * sizeof(WCHAR));
	TH->top += MultiByteToWideChar(CP_UTF8, 0, data, -1, (WCHAR *)ThLast(TH), size) * sizeof(WCHAR) - sizeof(WCHAR);
	return TRUE;
}

#if 0
/*===========================================================
 JS_GetOpaque()/JS_SetOpaque() を使わずにプロパティの１つとして
 実装する例。

 JSValueUnion 使用時のみ機能するように直接記載している。 */
// v = JS_MKPTR(JS_TAG_UNDEFINED, vptr) 相当。
#define Js_CtxPtr(/*JSValue*/ v, /*void*/ vptr) { (v)->u.ptr = vptr; (v)->tag = JS_TAG_UNDEFINED;}

#define info_index 0

#define SetQjsCtxPtr(ctx, this_obj, index, ptr) { JSValueConst val; Js_CtxPtr(&val, ptr); JS_SetPropertyUint32(ctx, this_obj, index, val); }

#define GetQjsCtxPtr(ctx, this_obj, index) (((JSValue)JS_GetPropertyUint32(ctx, this_obj, index)).u.ptr)
#endif

//=========================================================== エラー通知
static void SetObjectText(JSContext *ctx, ThSTRUCT *TH, JSValueConst val)
{
	const char *str;

	str = JS_ToCString(ctx, val);
	if ( str != NULL ) {
		THCatStringA(TH, str);
		THCatStringA(TH, "\r\n");
		JS_FreeCString(ctx, str);
	} else {
		THCatStringA(TH, "[exception]\r\n");
	}
}

void ShowScriptError(JSContext *ctx, PPXAPPINFOW *ppxa, PPXMCOMMANDSTRUCT *pxc)
{
	ThSTRUCT text = ThSTRUCT_InitData;
	JSValue val;
	BOOL is_error;
	JSValue exception_val;

	THCatStringW(&text, (const WCHAR *)L"-- PPxQjs.dll error --\r\n<root>: ");
	THCatStringW(&text, pxc->param);
	THCatStringW(&text, (const WCHAR *)L"\r\n");

	exception_val = JS_GetException(ctx);
	is_error = JS_IsError(ctx, exception_val);
	SetObjectText(ctx, &text, exception_val);
	if (is_error) {
		val = JS_GetPropertyStr(ctx, exception_val, "stack");
		if (!JS_IsUndefined(val)) {
			SetObjectText(ctx, &text, val);
		}
		JS_FreeValue(ctx, val);
	}
	JS_FreeValue(ctx, exception_val);
	PopupMessage(ppxa, (WCHAR *)text.bottom);
	THFree(&text);
}

//=========================================================== 常駐関係
DWORD_PTR USECDECL DummyPPxFunc(PPXAPPINFOW *ppxa, DWORD cmdID, void *uptr)
{
	return PPXA_INVALID_FUNCTION;
}

PPXAPPINFOW DummyPPxAppInfo = { DummyPPxFunc, L"", L"", NULL };

PPXMCOMMANDSTRUCT DummyPxc =
#ifndef _WIN64
	{L"", L"", 0, 0, NULL};
#else
	{L"", NULL, L"", 0, 0};
#endif

void SleepStayInstance(InstanceValueStruct *info, OldPPxInfoStruct *OldInfo)
{
	// appinfo が使えない状態で使われた時用のダミーを設定する。
	if ( OldInfo->ppxa != NULL ){
		info->ppxa = OldInfo->ppxa;
		info->pxc = OldInfo->pxc;
	}else{
		info->ppxa = (info->stay.ppxa != NULL ) ?
				info->stay.ppxa : &DummyPPxAppInfo;
		info->pxc = &DummyPxc;
	}
	info->stay.entry--;
}

// 常駐登録
void ChainStayInstance(InstanceValueStruct *info)
{
	PPXAPPINFOW *ppxa;

	// 情報用意
	ppxa = info->ppxa;
	info->stay.threadID = GetCurrentThreadId();
	info->stay.source = malloc((wcslen(info->pxc->param) + 1) * sizeof(WCHAR));
	wcscpy(info->stay.source, info->pxc->param);
	info->stay.hWnd = ppxa->hWnd;

	if ( PPxVersion == 0 ){
		InitializeCriticalSection(&StayLock);
		PPxVersion = (long)ppxa->Function(ppxa, PPXCMDID_VERSION, NULL);
	}

	// 登録
	EnterCriticalSection(&StayLock);

	Chain_FreeChainObject(StayChains, info->stay.chain, (void *)info);
	LeaveCriticalSection(&StayLock);

	if ( info->stay.ppxa == NULL ){
		PPXCMDID_NEWAPPINFO_STRUCT pns = {0, 0, NULL, NULL};
		PPXAPPINFOW *ppxa = info->ppxa;

		DWORD_PTR stayppxa = ppxa->Function(ppxa, PPXCMDID_NEWAPPINFO, &pns);
		if ( stayppxa != PPXA_INVALID_FUNCTION ){
			info->stay.ppxa = (PPXAPPINFOW *)stayppxa;
		}
	}

	// 通知を有効に
	ppxa->Function(ppxa, PPXCMDID_REQUIRE_CLOSETHREAD, 0);
}

// 常駐解除
void DropStayInstance(InstanceValueStruct *info)
{
	FREECHAIN *thischain = &info->stay.chain;
	FREECHAIN *chain = &StayChains;
	FREECHAIN *nextchain;

	EnterCriticalSection(&StayLock);
	for(;;){
		nextchain = chain->next;
		if ( nextchain == NULL ) break;
		if ( nextchain == thischain ){
			chain->next = nextchain->next;
			break;
		}
		chain = nextchain;
	}
	info->stay.threadID = 0;
	LeaveCriticalSection(&StayLock);
}

// 同じスレッドの常駐解放
void FreeStayInstance(PPXAPPINFOW *ppxa, BOOL terminate)
{
	InstanceValueStruct *chaininfo;
	DWORD ThreadID = GetCurrentThreadId();
	FREECHAIN *chain = &StayChains, *nextchain;

	for(;;){
		nextchain = chain->next;
		if ( nextchain == NULL ) break;

		chaininfo = (InstanceValueStruct *)nextchain->value;

		if ( (chaininfo->stay.threadID == ThreadID) &&
			 (terminate ||
			  (chaininfo->stay.hWnd == ppxa->hWnd)) ){
			chain->next = nextchain->next; // drop
			JSValue global = JS_GetGlobalObject(chaininfo->ctx);
			JSValue callfunc = JS_GetPropertyStr(chaininfo->ctx, global, "ppx_finally");
			if ( !JS_IsNull(callfunc) ){
				PPXAPPINFOW *oldppxa;
				JSValue result;

				oldppxa = chaininfo->stay.ppxa;
				chaininfo->stay.ppxa = ppxa;
				result = JS_Call(chaininfo->ctx, callfunc, global, 0, NULL);
				JS_FreeValue(chaininfo->ctx, result);
				chaininfo->stay.ppxa = oldppxa;
			}
			JS_FreeValue(chaininfo->ctx, callfunc);
			JS_FreeValue(chaininfo->ctx, global);

			CloseInstance(chaininfo);
			chain = &StayChains;
			continue;
		}
		chain = nextchain;
	}
}

//===========================================================
// Eval のソース情報を設定
int SetEvalInfo(JSContext *ctx, JSValueConst func_val, JS_BOOL is_main)
{
	JSModuleDef *module;
	char path[CMDLINESIZE];
	JSValue meta_obj;
	JSAtom module_name_atom;
	const char *module_name;

	module = JS_VALUE_GET_PTR(func_val);
	module_name_atom = JS_GetModuleName(ctx, module);
	module_name = JS_AtomToCString(ctx, module_name_atom);
	JS_FreeAtom(ctx, module_name_atom);
	if ( module_name == NULL ) return -1;
	strcpy(path, (strchr(module_name, ':') != NULL) ? "file://" : "");
	strcat(path, module_name);
	JS_FreeCString(ctx, module_name);

	meta_obj = JS_GetImportMeta(ctx, module);
	if ( JS_IsException(meta_obj) ) return -1;
	JS_DefinePropertyValueStr(ctx, meta_obj, "url", JS_NewString(ctx, path), JS_PROP_C_W_E);
	JS_DefinePropertyValueStr(ctx, meta_obj, "main", JS_NewBool(ctx, is_main), JS_PROP_C_W_E);
	JS_FreeValue(ctx, meta_obj);
	return 0;
}

char *LoadScript(const WCHAR *filename, char **TextImage, DWORD *sizeL)
{
	HANDLE hFile;
	int margin = 32;

	hFile = CreateFileW(filename, GENERIC_READ,
			FILE_SHARE_WRITE | FILE_SHARE_READ,
			NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if ( hFile == INVALID_HANDLE_VALUE ){
		return NULL;
	}
	*sizeL = GetFileSize(hFile, NULL);
	*TextImage = malloc(*sizeL + margin);
	if ( *TextImage == NULL ){
		CloseHandle(hFile);
		return NULL;
	}
	if ( FALSE == ReadFile(hFile, *TextImage, *sizeL, sizeL, NULL) ) *sizeL = 0;
	CloseHandle(hFile);
	memset(*TextImage + *sizeL, 0, margin);

	if ( memcmp(*TextImage, UTF8HEADER, UTF8HEADERSIZE) == 0 ){ // UTF-8
		*sizeL -= UTF8HEADERSIZE;
		return *TextImage + UTF8HEADERSIZE;
	} else{
		return *TextImage;
	}
}

char *GetScriptSource(PPXAPPINFOW *ppxa, const WCHAR *param, char **image, size_t *script_len, int file)
{
	char *script;

	if ( file == 0 ){
		*script_len = WideCharToMultiByte(CP_UTF8, 0, param, -1, NULL, 0, NULL, NULL) + 1;
		*image = script = malloc(*script_len);
		if ( script != NULL ){
			*script_len = WideCharToMultiByte(CP_UTF8, 0, param, -1, script, *script_len, NULL, NULL) - 1;
		}
	}else{
		DWORD sizeL;		// ファイルの大きさ

		script = LoadScript(param, image, &sizeL);
		if ( script == NULL ){
			PopupMessagef(ppxa, L"%s open error", param);
		}else{
			*script_len = (size_t)sizeL;
		}
	}
	return script;
}

int GetIntNumberW(const WCHAR *line)
{
	int n = 0;

	for ( ;; ){
		WCHAR code;

		code = *line;
		if ( (code < '0') || (code > '9') ) break;
		n = n * 10 + (WCHAR)((BYTE)code - (BYTE)'0');
		line++;
	}
	return n;
}

void CheckOption(PPXMCOMMANDSTRUCT *pxc, int *StayMode, char *InvokeName)
{
	WCHAR buf[256], *dest;

	pxc->param++;
	for (;;){
		dest = buf;
		for (;;){
			if ( *pxc->param == '\0' ) break;
			*dest++ = *pxc->param++;
			if ( *pxc->param != ',' ) continue;
			pxc->param++;
			break;
		}
		*dest = '\0';
		if ( buf[0] == '\0' ) break;
		if ( (buf[0] >= '0') && (buf[0] <= '9') ){
			*StayMode = GetIntNumberW(buf);
		}else{
			WideCharToMultiByte(CP_UTF8, 0, buf, -1, InvokeName, 256, NULL, NULL);
		}
	}
	pxc->param += wcslen(pxc->param) + 1; // ":～" をスキップ
	pxc->paramcount--;
}

#define RELESE
void CloseInstance(InstanceValueStruct *info)
{
	if ( info->stay.threadID != 0 ) DropStayInstance(info);
#ifndef RELESE
	printf("\r\nBefore FreeContext Disp %d/%d   Event %d/%d\r\n", CountDispRelease, CountDisp, CountEventRelease, CountEvent);
#endif
	ReleaseAllEvent(&info->EventChain);
	JS_FreeContext(info->ctx);
#ifndef RELESE
	printf("FreeContext Disp %d/%d   Event %d/%d\r\n", CountDispRelease, CountDisp, CountEventRelease, CountEvent);
#endif
	JS_FreeRuntime(info->rt);
#ifndef RELESE
	printf("FreeRuntime Disp %d/%d   Event %d/%d\r\n", CountDispRelease, CountDisp, CountEventRelease, CountEvent);
#endif
	if ( info->script_image != NULL ) free(info->script_image);
	if ( info->stay.ppxa != NULL ){
		HeapFree(ProcHeap, 0, info->stay.ppxa);
	}
	free(info);
}

int RunScript(PPXAPPINFOW *ppxa, PPXMCOMMANDSTRUCT *pxc, int file, BOOL module)
{
	JSValue result;
	InstanceValueStruct *info = NULL;
	BOOL error = FALSE;
	BOOL newctx = FALSE;
	size_t script_len;

	PPXMCOMMANDSTRUCT pxcbuf;
	int StayMode = ScriptStay_None;
	char InvokeName[256];
	ProcHeap = GetProcessHeap();
	OldPPxInfoStruct OldInfo;

	InvokeName[0] = '\0';
	OldInfo.ppxa = NULL;
	if ( pxc->paramcount > 0 ){
		if ( pxc->param[0] == ':' ){
			pxcbuf = *pxc;
			pxc = &pxcbuf;
			CheckOption(pxc, &StayMode, InvokeName);
		}else if ( (pxc->param[0] == '\0') && (pxc->paramcount >= 2) ){ // 第１パラメータが無いときは、shift
			// *script 変数(空 または :12345),source ができるようにする
			pxcbuf = *pxc;
			pxcbuf.paramcount--;
			pxcbuf.param += 1; // '\0' をスキップ
			pxc = &pxcbuf;
		}
	}

	if ( pxc->resultstring != NULL ) pxc->resultstring[0] = '\0';

	if ( StayChains.next != NULL ){
		FREECHAIN *chain = &StayChains, *nextchain;
		DWORD ThreadID = GetCurrentThreadId();
		HWND hWnd = ppxa->hWnd;
		const WCHAR *source = pxc->param;
		InstanceValueStruct *chaininfo;

		for(;;){
			nextchain = chain->next;
			if ( nextchain == NULL ) break;

			chaininfo = (InstanceValueStruct *)nextchain->value;

			if ( (chaininfo->stay.threadID == ThreadID) &&
				 (chaininfo->stay.hWnd == hWnd) &&
				 ((StayMode >= ScriptStay_Stay) ?
				   (chaininfo->stay.mode == StayMode) :
				   (wcscmp(chaininfo->stay.source, source) == 0)) ){
				info = chaininfo;
				info->stay.entry++;
				OldInfo.ppxa = info->ppxa;
				OldInfo.pxc = info->pxc;
				break;
			}
			chain = nextchain;
		}
	}

	if ( info == NULL ){
		info = malloc(sizeof(InstanceValueStruct));
		info->EventChain.next = NULL;
		info->script_image = NULL;
		info->stay.mode = StayMode;
		info->stay.threadID = 0;
		info->stay.entry = 0;
		info->stay.ppxa = NULL;

		info->rt = JS_NewRuntime();
		info->ctx = RunScriptNewContext(info->rt, info);
		JS_SetRuntimeOpaque(info->rt, NULL); /* fail safe */
		newctx = TRUE;
	}
	info->returncode = PPXMRESULT_DONE;
	info->quit = FALSE;
	info->ppxa = ppxa;
	info->pxc = pxc;

	if ( !newctx && (info->stay.mode >= ScriptStay_Stay) ){ // 常駐インスタンスの利用確認
		const WCHAR *param = pxc->param;
		int paramcount = pxc->paramcount;

		if ( InvokeName[0] == '\0' ){ // メソッド指定が無い
			if ( (StayMode < ScriptStay_Stay) || (paramcount == 0) || (param[0] == '\0') ){
					// インスタンス指定が無いなら ppx_resume
				strcpy(InvokeName, "ppx_resume");

				if ( paramcount > 0 ){ // source をスキップ
					param += wcslen(param) + 1; // 次のパラメータに
					paramcount--;
				}
			}else{ // インスタンス指定がある
				if ( file && (wcscmp(info->stay.source, pxc->param) == 0) ){
					// 常駐時スクリプトファイルなら ppx_resume
					strcpy(InvokeName, "ppx_resume");

					if ( paramcount > 0 ){ // source をスキップ
						param += wcslen(param) + 1; // 次のパラメータに
						paramcount--;
					}
				}else{ // 常駐時スクリプト以外なら eval
					char *image;
					char *script = GetScriptSource(ppxa, pxc->param, &image, &script_len, file);
					if ( script != NULL ){
						result = JS_Eval(info->ctx, script, script_len, "<2nd>", JS_EVAL_TYPE_GLOBAL);
						free(image);
					}
				}
			}
		}else{
			if ( StayMode < ScriptStay_Stay ){
				if ( paramcount > 0 ){
					param += wcslen(param) + 1; // source をスキップ
					paramcount--;
				}
			}
		}

		if ( InvokeName[0] != '\0' ){ // メソッド実行
			JSValue global = JS_GetGlobalObject(info->ctx);
			JSValue callfunc = JS_GetPropertyStr(info->ctx, global, InvokeName );
			if ( !JS_IsNull(callfunc) ){
				JSValue *argv = malloc(sizeof(JSValue) * paramcount);
				for ( int i = 0; i < paramcount; i++ ){
					argv[i] = JS_NewStringW(info->ctx, param);
					param += wcslen(param) + 1;
				}
				result = JS_Call(info->ctx, callfunc, global, paramcount, argv);
				if ( JS_IsException(result) == FALSE ){
					PPxSetResult(info->ctx, JS_NULL, result);
				}
				for ( unsigned int i = 0 ; i < paramcount; i++ ){
					JS_FreeValue(info->ctx, argv[i]);
				}
				free(argv);
			}else{
				PopupMessage(ppxa, L"not found function");
				result = JS_NULL;
			}
			JS_FreeValue(info->ctx, callfunc);
			JS_FreeValue(info->ctx, global);
		}
	}else{
		if ( InvokeName[0] != '\0' ){
			PopupMessage(ppxa, L"not found instance");
			result = JS_NULL;
		}else if ( pxc->paramcount == 0 ){
			PopupMessage(ppxa, L"no source parameter");
			result = JS_NULL;
		}else{
			char *image;
			char *script = GetScriptSource(ppxa, pxc->param, &image, &script_len, file);
			if ( script != NULL ){
				if ( info->script_image == NULL ) info->script_image = image;
				if ( module == FALSE ){
					result = JS_Eval(info->ctx, script, script_len, "<root>", JS_EVAL_TYPE_GLOBAL);
				}else{
					result = JS_Eval(info->ctx, script, script_len, "<root>", JS_EVAL_TYPE_MODULE | JS_EVAL_FLAG_COMPILE_ONLY);
					if ( JS_IsException(result) == FALSE ){
						SetEvalInfo(info->ctx, result, TRUE); // JS_EvalFunction の情報
						JSValue newresult = JS_EvalFunction(info->ctx, result);
						JS_FreeValue(info->ctx, result);
						result = newresult;
					}
					if ( image != info->script_image ) free(image);
				}
			}else{
				result = JS_NULL;
			}
		}
	}

	if ( JS_IsError(info->ctx, result) || JS_IsException(result) ){
		if ( !info->quit ) ShowScriptError(info->ctx, ppxa, pxc);
		error = TRUE;
	}
	JS_FreeValue(info->ctx, result);
	int returncode = info->returncode;

	if ( !error && (info->stay.mode >= ScriptStay_Stay) ){ // 常駐時
		if ( info->stay.threadID == 0 ){ // 未登録なので登録
			ChainStayInstance(info);
		}
		SleepStayInstance(info, &OldInfo);
	}else{ // 非常駐時
		info->stay.entry--;
		CloseInstance(info);
	}
	return returncode;
}

//=========================================================== 型変換
JSValue JS_NewStringW(JSContext *ctx, const WCHAR *str)
{
	char bufA[CMDLINESIZE];

	if ( str == NULL ) return JS_NULL;
	if ( str[0] == '\0' ) return JS_NewString(ctx, "");

	int script_len = WideCharToMultiByte(CP_UTF8, 0, str, -1, bufA, sizeof(bufA), NULL, NULL);
	if ( script_len > 0 ) return JS_NewString(ctx, bufA);

	script_len = WideCharToMultiByte(CP_UTF8, 0, str, -1, NULL, 0, NULL, NULL) + 8;
	char *script = malloc(script_len);
	if ( script == NULL ){
		JS_ThrowOutOfMemory(ctx);
		return JS_EXCEPTION;
	}
	WideCharToMultiByte(CP_UTF8, 0, str, -1, script, script_len, NULL, NULL);

	JSValue val = JS_NewString(ctx, script);
	free(script);
	return val;
}

WCHAR *GetJsLongString(JSContext *ctx, JSValueConst js_var, WCHAR *bufW)
{
	const char *varA;
	size_t lenA;

	varA = JS_ToCStringLen(ctx, &lenA, js_var);
	if ( varA != NULL ){
		int lenW;

		lenW = MultiByteToWideChar(CP_UTF8, 0, varA, lenA + 1, bufW, CMDLINESIZE);
		if ( lenW > 0 ){
			JS_FreeCString(ctx, varA);
			return bufW;
		}
		lenW = MultiByteToWideChar(CP_UTF8, 0, varA, lenA + 1, NULL, 0);
		WCHAR *valW = malloc(lenW * sizeof(WCHAR));
		if ( valW != NULL ){
			MultiByteToWideChar(CP_UTF8, 0, varA, lenA + 1, valW, lenW);
			JS_FreeCString(ctx, varA);
			return valW;
		}
		JS_ThrowOutOfMemory(ctx);
		JS_FreeCString(ctx, varA);
	}
	// 失敗
	bufW[0] = '\0';
	return bufW;
}

const TCHAR InfoForm[] = L" %d";

int StayInfo(PPXAPPINFOW *ppxa, PPXMCOMMANDSTRUCT *pxc)
{
	InstanceValueStruct *chaininfo;
	FREECHAIN *chain = &StayChains, *nextchain;
	DWORD ThreadID = GetCurrentThreadId();
	HWND hWnd = ppxa->hWnd;
	WCHAR *dest = pxc->resultstring, *destmax = dest + (CMDLINESIZE - 12);
	int StayMode = ScriptStay_None;

	dest[0] = '\0';

	if ( pxc->paramcount > 0 ){
		const WCHAR *source = pxc->param;

		if ( *source == ':' ) source++;
		StayMode = GetIntNumberW(source);
		if ( StayMode >= ScriptStay_Stay ){
			dest[0] = '0';
			dest[1] = '\0';
		}
	}

	if ( StayChains.next == NULL ) return PPXMRESULT_DONE;

	for(;;){
		nextchain = chain->next;
		if ( nextchain == NULL ) break;

		chaininfo = (InstanceValueStruct *)nextchain->value;

		if ( (chaininfo->stay.threadID == ThreadID) &&
			 (chaininfo->stay.hWnd == hWnd) ){
			// インスタンス指定有り→個別結果
			if ( StayMode >= ScriptStay_Stay ){
				if ( chaininfo->stay.mode == StayMode ){
					dest[0] = '1';
					break;
				}
			}else{ // 指定無し→一覧
				dest += wsprintf(dest, (dest == pxc->resultstring) ?
						InfoForm + 1 : InfoForm,
						chaininfo->stay.mode);
				if ( dest >= destmax ) break;
			}
		}
		chain = nextchain;
	}
	return PPXMRESULT_DONE;
}
