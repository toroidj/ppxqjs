/*-----------------------------------------------------------------------------
	Paper Plane xUI QuickJS Script Module							(c)TORO
	OLE オブジェクトの代替物 ADODB.stream / Scripting.Filesystemobject
-----------------------------------------------------------------------------*/
#define STRICT
#define UNICODE
#include <stddef.h>
#include <windows.h>
#include <dispex.h>
#include <combaseapi.h>
#include "quickjs.h"
#include "TOROWIN.H"
#include "PPCOMMON.H"
#include "ppxqjs.h"

JSClassID ClassID_Folder = 0;

//================================================= CreateObject - ADODB.stream
static JSValue StreamClose(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JS_SetPropertyStr(ctx, this_obj, "DATA", JS_NULL);
	return JS_UNDEFINED;
}

static JSValue StreamEOS(JSContext *ctx, JSValueConst this_obj)
{
	int pos, size;
	JSValue jsvar;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "Position");
	JS_ToInt32(ctx, &pos, jsvar);
	JS_FreeValue(ctx, jsvar);
	jsvar = JS_GetPropertyStr(ctx, this_obj, "Size");
	JS_ToInt32(ctx, &size, jsvar);
	JS_FreeValue(ctx, jsvar);
	return JS_NewBool(ctx, (pos < size));
}

#define CP__AUTODETECT 1

UINT StreamCodePage(JSContext *ctx, JSValueConst this_obj)
{
	UINT codepage = CP__AUTODETECT;
	const char *charset;
	JSValue jsvar;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "CharSet");
	charset = JS_ToCString(ctx, jsvar);
	JS_FreeValue(ctx, jsvar);

	if ( strstr(charset, "auto") != 0 ){
		codepage = CP__AUTODETECT;
	}else if ( (stricmp(charset, "utf-8") == 0) || (stricmp(charset, "utf8") == 0) ){
		codepage = CP_UTF8;
	}else if ( (stricmp(charset, "unicode") == 0) || (stricmp(charset, "utf-16le") == 0) || (stricmp(charset, "utf16le") == 0) ){
		codepage = CP__UTF16L;
	}else{
		codepage = CP_ACP;
	}
	JS_FreeCString(ctx, charset);
	return codepage;
}

JSValue StreamLoadFromFile(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR filename[CMDLINESIZE];
	HANDLE hFile;
	DWORD sizeL;
	char *TextImage;
	UINT codepage;

	GetJsShortString(ctx, argv[0], filename);
	hFile = CreateFileW(filename, GENERIC_READ,
			FILE_SHARE_WRITE | FILE_SHARE_READ,
			NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if ( hFile == INVALID_HANDLE_VALUE ){
		return JS_NULL;
	}
	sizeL = GetFileSize(hFile, NULL);
	TextImage = static_cast<char *>(malloc(sizeL + 1));
	if ( TextImage == NULL ){
		CloseHandle(hFile);
		return JS_NULL;
	}
	if ( FALSE == ReadFile(hFile, TextImage, sizeL, &sizeL, NULL) ) sizeL = 0;
	CloseHandle(hFile);
	TextImage[sizeL] = '\0';

	codepage = StreamCodePage(ctx, this_obj);
	if ( codepage == CP__AUTODETECT ){ // 自動判別
		if ( ((sizeL & 1) == 0) && (sizeL > 2) && (TextImage[1] == '\0') ){
			codepage = CP__UTF16L;
		}else if ( MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, TextImage, -1, NULL, 0) > 0 ){
			codepage = CP_UTF8;
		}else{
			codepage = CP_ACP;
		}
		JS_SetPropertyStr(ctx, this_obj, "CharSet", JS_NewString(ctx,
				(codepage == CP_UTF8) ? "utf-8" : (codepage == CP__UTF16L) ? "unicode" : "cp_acp"));
	}

	if ( codepage != CP_UTF8 ){
		WCHAR *textW;
		if ( codepage == CP__UTF16L ){ // unicode
			textW = (WCHAR *)TextImage;
		}else{ // CP_ACP等 shift_jis(ansi)
			int size = MultiByteToWideChar(codepage, 0, TextImage, -1, NULL, 0);
			textW = static_cast<WCHAR *>(malloc(size * sizeof(WCHAR)));
			size = MultiByteToWideChar(codepage, 0, TextImage, -1, textW, size);
			free(TextImage);
		}
		int sizeL = WideCharToMultiByte(CP_UTF8, 0, textW, -1, NULL, 0, NULL, NULL) + 1;
		TextImage = static_cast<char *>(malloc(sizeL * sizeof(WCHAR)));
		sizeL = WideCharToMultiByte(CP_UTF8, 0, textW, -1, TextImage, sizeL, NULL, NULL) - 1;
		free(textW);
	}

	JS_SetPropertyStr(ctx, this_obj, "DATA", JS_NewString(ctx, TextImage));
	JS_SetPropertyStr(ctx, this_obj, "Position", JS_NewInt32(ctx, 0));
	JS_SetPropertyStr(ctx, this_obj, "Size", JS_NewInt32(ctx, sizeL));
	free(TextImage);
	return JS_UNDEFINED;
}

JSValue StreamDummyFunc(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return JS_UNDEFINED;
}

JSValue StreamRead(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	int pos, size, type = -1;
	JSValue jsvar;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "Position");
	JS_ToInt32(ctx, &pos, jsvar);
	JS_FreeValue(ctx, jsvar);
	jsvar = JS_GetPropertyStr(ctx, this_obj, "Size");
	JS_ToInt32(ctx, &size, jsvar);
	JS_FreeValue(ctx, jsvar);
	if ( argc > 0 ) JS_ToInt32(ctx, &type, argv[0]);

	if ( pos < 0 ) pos = 0;
	if ( pos >= size ) return JS_NULL;

	JSValue val = JS_GetPropertyStr(ctx, this_obj, "DATA");
	if ( (type == -1) && (pos == 0) ){ // 全体を一度に読み込む
		pos = size;
	}else{
		const char *strorg, *strptr;
		strorg = strptr = JS_ToCString(ctx, val);
		JS_FreeValue(ctx, val);
		strptr += pos;
		if ( type == -1 ){
			pos = size;
			val = JS_NewString(ctx, strptr);
		}else if ( type == -2 ){
			const char *ptr = strptr;
			int lf = 0;
			for (;;){
				char chr;
				chr = *ptr;
				if ( chr == '\0' ) break;
				ptr++;
				if ( chr == '\r' ){
					lf = 1;
					if ( *ptr == '\n' ){
						ptr++;
						lf = 2;
					}
					break;
				}
				if ( chr == '\n' ){
					lf = 1;
					break;
				}
			}
			type = ptr - strptr;
			val = JS_NewStringLen(ctx, strptr, type - lf);
			pos += type;
		}else if ( type >= 0 ){
			val = JS_NewStringLen(ctx, strptr, type);
			pos += type;
		}else{
			val = JS_NULL;
		}
		JS_FreeCString(ctx, strorg);
	}
	JS_SetPropertyStr(ctx, this_obj, "Position", JS_NewInt32(ctx, pos));
	return val;
}

static JSValue StreamSkipLine(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	JSValueConst argv1 = JS_NewInt32(ctx, -2);
	JSValue var = StreamRead(ctx, this_obj, 1, &argv1);
	JS_FreeValue(ctx, var);
	return JS_UNDEFINED;
}

JSValue StreamSaveToFile(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR filename[CMDLINESIZE];
	HANDLE hFile;
	DWORD sizeL;
	UINT codepage;

	GetJsShortString(ctx, argv[0], filename);
	hFile = CreateFileW(filename, GENERIC_WRITE,
			FILE_SHARE_WRITE | FILE_SHARE_READ,
			NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if ( hFile == INVALID_HANDLE_VALUE ) return JS_NULL;

	JSValue val = JS_GetPropertyStr(ctx, this_obj, "DATA");
	if ( JS_IsString(val) ){
		const char *TextImage = JS_ToCString(ctx, val);

		if ( TextImage != NULL ){
			codepage = StreamCodePage(ctx, this_obj);
			if ( codepage == CP__AUTODETECT ) codepage = CP_UTF8;
			if ( codepage != CP_UTF8 ){
				WCHAR *textW;
				DWORD sizeL;
				int sizeW = MultiByteToWideChar(CP_UTF8, 0, TextImage, -1, NULL, 0);

				textW = static_cast<WCHAR *>(malloc(sizeW * sizeof(WCHAR)));
				sizeW = MultiByteToWideChar(CP_UTF8, 0, TextImage, -1, textW, sizeW);
				if ( codepage == CP__UTF16L ){ // unicode
					WriteFile(hFile, textW, (sizeW - 1) * sizeof(WCHAR), &sizeL, NULL);
				}else{ // CP_ACP等 shift_jis(ansi)
					char *textA;
					int sizeA = WideCharToMultiByte(codepage, 0, textW, -1, NULL, 0, NULL, NULL) + 1;
					textA = static_cast<char *>(malloc(sizeA));
					sizeA = WideCharToMultiByte(codepage, 0, textW, -1, textA, sizeA, NULL, NULL);
					WriteFile(hFile, textA, sizeA - 1, &sizeL, NULL);
					free(textA);
					free(textW);
				}
			}else{
				WriteFile(hFile, TextImage, strlen(TextImage), &sizeL, NULL);
			}
			JS_FreeCString(ctx, TextImage);
		}
	}
	JS_FreeValue(ctx, val);
	CloseHandle(hFile);
	return JS_UNDEFINED;
}

JSValue StreamWrite(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	int pos, size, addsize = 0, option = 0, lfcode;
	JSValue jsvar;
	const char *adddata = NULL;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "Position");
	JS_ToInt32(ctx, &pos, jsvar);
	JS_FreeValue(ctx, jsvar);

	jsvar = JS_GetPropertyStr(ctx, this_obj, "Size");
	JS_ToInt32(ctx, &size, jsvar);
	JS_FreeValue(ctx, jsvar);

	if ( argc >= 1 ){
		if ( argc >= 2 ){
			JS_ToInt32(ctx, &option, argv[1]);
			if ( option != 0 ){
				lfcode = -1;
				jsvar = JS_GetPropertyStr(ctx, this_obj, "LineSeparator");
				JS_ToInt32(ctx, &lfcode, jsvar);
				JS_FreeValue(ctx, jsvar);
				if ( lfcode < 0 ){
					option = 2; // \r\n 分
				}else{
					option = 1;
				}
			}
		}
		adddata = JS_ToCString(ctx, argv[0]);
		if ( adddata != NULL ) addsize = strlen(adddata);
	}

	if ( (pos == 0) && (option == 0) ){
		if ( addsize > 0 ){
			JS_SetPropertyStr(ctx, this_obj, "DATA", JS_DupValue(ctx, argv[0]));
			pos = size = addsize;
		}
	}else{
		const char *data;
		char *newdata;
		int filesize;

		filesize = (size > (pos + addsize + option)) ? size : (pos + addsize + option);
		newdata = static_cast<char *>(malloc(filesize + 1));
		if ( pos > size ) memset(newdata + size, 0, pos - size);

		JSValue jsvar = JS_GetPropertyStr(ctx, this_obj, "DATA");
		data = JS_ToCString(ctx, jsvar);

		memcpy(newdata, data, size);
		JS_FreeCString(ctx, data);
		JS_FreeValue(ctx, jsvar);

		if ( addsize > 0 ){
			memcpy(newdata + pos, adddata, addsize);
			pos += addsize;
			if ( option ){
				if ( lfcode < 0 ){
					newdata[pos] = '\r';
					newdata[pos + 1] = '\n';
					pos += 2;
				}else{
					newdata[pos++] = lfcode;
				}
			}
		}else if ( argc == 0 ){ // SetEOS
			filesize = size = pos;
		}
		JS_SetPropertyStr(ctx, this_obj, "DATA", JS_NewStringLen(ctx, newdata, filesize));
		if ( pos > size ) size = pos;
	}
	if ( adddata != NULL ) JS_FreeCString(ctx, adddata);
	JS_SetPropertyStr(ctx, this_obj, "Position", JS_NewInt32(ctx, pos));
	JS_SetPropertyStr(ctx, this_obj, "Size", JS_NewInt32(ctx, size));
	return JS_UNDEFINED;
}

static JSValue StreamSetEOS(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return StreamWrite(ctx, this_obj, 0, NULL);
}

static const JSCFunctionListEntry StreamFunctionList[] = {
//	JS_CGETSET_DEF("AtEndOfLine", StreamEOS, NULL), // textstream
//	JS_CGETSET_DEF("AtEndOfStream", StreamEOS, NULL), // textstream
//	JS_CGETSET_DEF("Column", StreamEOS, NULL), // textstream
//	JS_CGETSET_DEF("Line", StreamEOS, NULL), // textstream
	JS_CFUNC_DEF("Cancel", 0, StreamDummyFunc), // adodb
	JS_PROP_STRING_DEF("CharSet", "utf-8", JS_PROP_WRITABLE), // adodb // _autodetect, _autodetect_all, utf-8 unicode
	JS_CFUNC_DEF("Close", 0, StreamClose), // textstream adodb
	JS_PROP_INT32_DEF("CopyTo", 0, 0), // adodb CopyTo(DestStream, NumChars);
	JS_CGETSET_DEF("EOS", StreamEOS, NULL), // adodb
	JS_CFUNC_DEF("Flush", 0, StreamDummyFunc), // adodb
	JS_PROP_INT32_DEF("LineSeparator", -1, JS_PROP_WRITABLE), // adodb
	JS_CFUNC_DEF("LoadFromFile", 1, StreamLoadFromFile), // adodb
	JS_PROP_INT32_DEF("Mode", 0, JS_PROP_WRITABLE), // adodb
	JS_CFUNC_DEF("Open", 0, StreamDummyFunc), // adodb
	JS_PROP_INT32_DEF("Position", 0, JS_PROP_WRITABLE), // adodb(64)
	JS_CFUNC_DEF("Read", 1, StreamRead), // adodb/textstream
//	JS_CFUNC_DEF("ReadAll", 1, StreamRead), // textstream
//	JS_CFUNC_DEF("ReadLine", 1, StreamRead), // textstream
	JS_CFUNC_DEF("ReadText", 1, StreamRead), // adodb
	JS_CFUNC_DEF("SaveToFile", 1, StreamSaveToFile), // adodb
	JS_CFUNC_DEF("SetEOS", 0, StreamSetEOS), // adodb
	JS_PROP_INT32_DEF("Size", 0, JS_PROP_WRITABLE), // adodb(getのみ)
//	JS_CFUNC_DEF("Skip", 1, StreamRead), // textstream
	JS_CFUNC_DEF("SkipLine", 0, StreamSkipLine), // adodb, textstream
	JS_CFUNC_DEF("Stat", 1, StreamDummyFunc), // adodb
	JS_PROP_INT32_DEF("State", 1, 0), // adodb
	JS_PROP_INT32_DEF("Type", 2, JS_PROP_WRITABLE), // adodb
	JS_CFUNC_DEF("Write", 1, StreamWrite), // adodb/textstream
//	JS_CFUNC_DEF("WriteBlankLines", 1, StreamWrite), // textstream
//	JS_CFUNC_DEF("WriteLine", 1, StreamWrite), // textstream
	JS_CFUNC_DEF("WriteText", 1, StreamWrite), // adodb
	JS_PROP_STRING_DEF("DATA", "", JS_PROP_WRITABLE), // Local data
};

JSValue New_ADODB_Stream(JSContext *ctx)
{
	JSValue newobj = JS_NewObject(ctx);
	JS_SetPropertyFunctionList(ctx, newobj, StreamFunctionList, _countof(StreamFunctionList));
	JS_PreventExtensions(ctx, newobj);
	return newobj;
}

//===================== CreateObject Scripting.Filesystemobject - folder & file
const char *FSOTimeName[3] = {"_DateCreated", "_DateLastAccessed", "_DateLastModified"};

static JSValue GetFolderTime(JSContext *ctx, JSValueConst this_obj, int magic)
{
	JSValue jsvar;
	int64_t utctime;
	jsvar = JS_GetPropertyStr(ctx, this_obj, FSOTimeName[magic]);
	JS_ToInt64(ctx, &utctime, jsvar);
	JS_FreeValue(ctx, jsvar);
	return JS_NewInt64(ctx, (utctime - 116444736000000000) / 10000);
/*
	char buf[40];
	size_t len;
	len = sprintf(buf, "new Date(%lld);", (utctime - 116444736000000000) / 10000);
	return JS_Eval(ctx, buf, len, "<>", 0);
*/
}

JSValue GetFolderParentFolder(JSContext *ctx, JSValueConst this_obj)
{
	WCHAR path[CMDLINESIZE];
	WCHAR cmd[CMDLINESIZE];
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	JSValue jsvar;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "Path");
	GetJsShortString(ctx, jsvar, path);
	JS_FreeValue(ctx, jsvar);

	wsprintfW(cmd, L"%%*name(D,%%(%s%%))", path);
	if ( ppxa->Function(ppxa, PPXCMDID_EXTRACT, cmd) == PPXA_NO_ERROR ){
		return JS_NewStringW(ctx, cmd);
	}else{
		return JS_NewString(ctx, "");
	}
}

static JSValue GetFolderEntries(JSContext *ctx, JSValueConst this_obj, int magic);
static const JSCFunctionListEntry FolderFunctionList[] = {
	JS_PROP_INT32_DEF("Attributes", 0, JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
	JS_CGETSET_MAGIC_DEF("DateCreated", GetFolderTime, NULL, 0),
	JS_CGETSET_MAGIC_DEF("DateLastAccessed", GetFolderTime, NULL, 1),
	JS_CGETSET_MAGIC_DEF("DateLastModified", GetFolderTime, NULL, 2),
//	JS_PROP_INT32_DEF("Add", 0, 0), // folder のみ
//	JS_PROP_INT32_DEF("Copy", 0, 0),
//	JS_PROP_INT32_DEF("Delete", 0, 0),
//	JS_PROP_INT32_DEF("Move", 0, 0),
//	JS_PROP_INT32_DEF("OpenAsTextStream", 0, 0), // file のみ
//	JS_PROP_INT32_DEF("CreateTextFile", 0, 0),  // folder のみ
//	JS_PROP_INT32_DEF("Drive", 0, 0),
	JS_CGETSET_MAGIC_DEF("Files", GetFolderEntries, NULL, 0),
//	JS_PROP_INT32_DEF("IsRootFolder", 0, 0), // folder のみ
	JS_PROP_STRING_DEF("Name", "", JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
	JS_PROP_STRING_DEF("Path", "", JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
	JS_CGETSET_DEF("ParentFolder", GetFolderParentFolder, NULL),
//	JS_PROP_INT32_DEF("ShortName", 0, 0),
//	JS_PROP_INT32_DEF("ShortPath", 0, 0),
	JS_PROP_INT64_DEF("_DateCreated", 0, JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
	JS_PROP_INT64_DEF("_DateLastAccessed", 0, JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
	JS_PROP_INT64_DEF("_DateLastModified", 0, JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
	JS_PROP_INT64_DEF("Size", 0, JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
	JS_CGETSET_MAGIC_DEF("SubFolders", GetFolderEntries, NULL, 1), // folder のみ
	JS_PROP_INT32_DEF("Type", 0, JS_PROP_WRITABLE | JS_PROP_CONFIGURABLE),
};

JSClassDef FolderClassDef = {
	"_FoldeR_", // class_name
	NULL, // finalizer
	NULL, // gc_mark
	NULL, // call
	NULL // exotic
};

JSValue MakeFolderItem(JSContext *ctx, WIN32_FIND_DATAW *ff)
{
	JSValue newobj;
	JSRuntime *rt = JS_GetRuntime(ctx);

	if ( ClassID_Folder == 0 ) JS_NewClassID(&ClassID_Folder);
	if ( !JS_IsRegisteredClass(rt, ClassID_Folder) ){ // クラスを登録
		JS_NewClass(rt, ClassID_Folder, &FolderClassDef);
		JSValue ProtoClass = JS_NewObject(ctx);
		JS_SetPropertyFunctionList(ctx, ProtoClass, FolderFunctionList, _countof(FolderFunctionList));
		JS_SetClassProto(ctx, ClassID_Folder, ProtoClass);
	}
	newobj = JS_NewObjectClass(ctx, ClassID_Folder);
	JS_SetPropertyStr(ctx, newobj, "Attributes", JS_NewInt32(ctx, ff->dwFileAttributes));
	JS_SetPropertyStr(ctx, newobj, "Name", JS_NewStringW(ctx, ff->cFileName));
	JS_SetPropertyStr(ctx, newobj, "Size", JS_NewInt64(ctx, (__int64)ff->nFileSizeLow + ((__int64)ff->nFileSizeHigh << 32)));
	JS_SetPropertyStr(ctx, newobj, "_DateCreated", JS_NewInt64(ctx, (__int64)ff->ftCreationTime.dwLowDateTime + ((__int64)ff->ftCreationTime.dwHighDateTime << 32)));
	JS_SetPropertyStr(ctx, newobj, "_DateLastAccessed", JS_NewInt64(ctx, (__int64)ff->ftLastAccessTime.dwLowDateTime + ((__int64)ff->ftLastAccessTime.dwHighDateTime << 32)));
	JS_SetPropertyStr(ctx, newobj, "_DateLastModified", JS_NewInt64(ctx, (__int64)ff->ftLastWriteTime.dwLowDateTime + ((__int64)ff->ftLastWriteTime.dwHighDateTime << 32)));
	return newobj;
}

//===================================== CreateObject Scripting.Filesystemobject
static JSValue GetFolderEntries(JSContext *ctx, JSValueConst this_obj, int magic)
{
	JSValue jsvar;
	WCHAR path[CMDLINESIZE];
	WCHAR filename[CMDLINESIZE];
	int index = 0;
	WIN32_FIND_DATAW ff;
	HANDLE hFF;

	jsvar = JS_GetPropertyStr(ctx, this_obj, "Path");
	GetJsShortString(ctx, jsvar, path);
	JS_FreeValue(ctx, jsvar);

	wsprintfW(filename, L"%s\\*", path);
	hFF = FindFirstFileW(filename, &ff);
	if ( hFF == INVALID_HANDLE_VALUE ) return JS_NULL;

	JSValue newobj = JS_NewArray(ctx);
	do{
		JSValue jsitem;

		if ( IsRelativeDir(ff.cFileName) ) continue;
		if ( magic == 0 ){ // files
			if ( ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY ) continue;
		}else{ // subfolders
			if ( !(ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ) continue;
		}
		jsitem = MakeFolderItem(ctx, &ff);
		wsprintfW(filename, L"%s\\%s", path, ff.cFileName);
		JS_SetPropertyStr(ctx, jsitem, "Path", JS_NewStringW(ctx, filename));
		JS_SetPropertyUint32(ctx, newobj, index++, jsitem);
	}while(FindNextFile(hFF, &ff));
	FindClose(hFF);
	return newobj;
}

JSValue FsoGetEntry(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR argvpath[CMDLINESIZE];
	WCHAR path[CMDLINESIZE];
	WIN32_FIND_DATAW ff;
	HANDLE hFF;
	JSValue newobj;

	GetJsShortString(ctx, argv[0], argvpath);
	GetFullPathName(argvpath, CMDLINESIZE, path, NULL);

	hFF = FindFirstFileW(path, &ff);
	if ( hFF == INVALID_HANDLE_VALUE ) return JS_NULL;
	FindClose(hFF);
	newobj = MakeFolderItem(ctx, &ff);
	JS_SetPropertyStr(ctx, newobj, "Path", JS_NewStringW(ctx, path));
	return newobj;
}

void FsoExec2param(JSContext *ctx, int argc, JSValueConst *argv, const WCHAR *cmdline)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR src[CMDLINESIZE], dest[CMDLINESIZE], cmd[CMDLINESIZE * 2 + 128];

	if ( argc < 2 ) return;
	GetJsShortString(ctx, argv[0], src);
	GetJsShortString(ctx, argv[1], dest);
	wsprintfW(cmd, cmdline, src, dest);
	ppxa->Function(ppxa, PPXCMDID_EXECUTE, cmd);
}

JSValue FsoBuildPath(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	char *path, pathbuf[CMDLINESIZE];
	const char *str;

	str = JS_ToCString(ctx, argv[0]);
	strcpy(pathbuf, str);
	JS_FreeCString(ctx, str);
	path = pathbuf + strlen(pathbuf);
	if ( (path != pathbuf) && (*(path - 1) != '\\') && (*(path - 1) != '/') ){
		*path++ = '\\';
	}
	str = JS_ToCString(ctx, argv[1]);
	strcpy(path, str);
	JS_FreeCString(ctx, str);
	return JS_NewString(ctx, pathbuf);
}

JSValue FsoCopyFile(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	FsoExec2param(ctx, argc, argv,
			L"%%Obn cmd.exe /c copy /Y %%(\"%s\" \"%s\"%%)");
	return JS_UNDEFINED;
}

JSValue FsoCopyFolder(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	FsoExec2param(ctx, argc, argv,
			L"%%Obn xcopy.exe /Y %%(\"%s\" \"%s\"%%)");
	return JS_UNDEFINED;
}

JSValue FsoMoveFile(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	FsoExec2param(ctx, argc, argv,
			L"%%Obn cmd.exe /c move /Y %%(\"%s\" \"%s\"%%)");
	return JS_UNDEFINED;
}

JSValue FsoMoveFolder(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return JS_UNDEFINED;
}

JSValue FsoCreateFolder(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR bufw[CMDLINESIZE];

	GetJsShortString(ctx, argv[0], bufw);
	return JS_NewBool(ctx, CreateDirectoryW(bufw, NULL));
}

JSValue FsoDeleteFile(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR bufw[CMDLINESIZE];

	GetJsShortString(ctx, argv[0], bufw);
	return JS_NewBool(ctx, DeleteFileW(bufw));
}

JSValue FsoDeleteFolder(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR src[CMDLINESIZE], cmd[CMDLINESIZE * 2 + 128];

	GetJsShortString(ctx, argv[0], src);
	wsprintfW(cmd, L"%%Obn cmd.exe /c rd /S /Q %%(\"%s\"%%)", src);
	ppxa->Function(ppxa, PPXCMDID_EXECUTE, cmd);
	return JS_UNDEFINED;
}

JSValue FsoFileExists(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR bufw[CMDLINESIZE];
	DWORD attr;

	GetJsShortString(ctx, argv[0], bufw);
	attr = GetFileAttributesW(bufw);

	return JS_NewBool(ctx, (attr != MAX32) && !(attr & FILE_ATTRIBUTE_DIRECTORY));
}

JSValue FsoFolderExists(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	WCHAR bufw[CMDLINESIZE];
	DWORD attr;

	GetJsShortString(ctx, argv[0], bufw);
	attr = GetFileAttributesW(bufw);

	return JS_NewBool(ctx, (attr != MAX32) && (attr & FILE_ATTRIBUTE_DIRECTORY));
}

int TempCount = 0;

JSValue FsoGetTempName(JSContext *ctx, JSValueConst this_obj)
{
	char buf[32];
	sprintf(buf, "rad%05X.tmp", (unsigned int)((GetTickCount() & 0xffff0) + (TempCount++ & 0xf)));
	return JS_NewString(ctx, buf);
}

static JSValue FsoExtract1param(JSContext *ctx, int argc, JSValueConst *argv, const WCHAR *cmdline)
{
	PPXAPPINFOW *ppxa = GetCtxPpxa(ctx);
	WCHAR src[CMDLINESIZE];
	WCHAR cmd[CMDLINESIZE];

	if ( argc < 1 ) return JS_NULL;
	GetJsShortString(ctx, argv[0], src);
	wsprintfW(cmd, cmdline, src);
	if ( ppxa->Function(ppxa, PPXCMDID_EXTRACT, cmd) == PPXA_NO_ERROR ){
		return JS_NewStringW(ctx, cmd);
	}else{
		return JS_NewString(ctx, "");
	}
}

static JSValue FsoGetAbsolutePathName(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return FsoExtract1param(ctx, argc, argv, L"%%*name(DC,%%(%s%%))");
}

static JSValue FsoGetBaseName(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return FsoExtract1param(ctx, argc, argv, L"%%*name(X,%%(%s%%))");
}

static JSValue FsoGetDriveName(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return FsoExtract1param(ctx, argc, argv, L"%%*name(H,%%(%s%%))");
}

static JSValue FsoGetExtensionName(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return FsoExtract1param(ctx, argc, argv, L"%%*name(T,%%(%s%%))");
}

static JSValue FsoGetFileName(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return FsoExtract1param(ctx, argc, argv, L"%%*name(C,%%(%s%%))");
}

static JSValue FsoGetParentFolderName(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	return FsoExtract1param(ctx, argc, argv, L"%%*name(D,%%(%s%%))");
}

static JSValue FsoGetSpecialFolder(JSContext *ctx, JSValueConst this_obj, int argc, JSValueConst *argv)
{
	int index;
	WCHAR bufW[CMDLINESIZE];

	JS_ToInt32(ctx, &index, argv[0]);
	switch (index){
		case 0:
			GetWindowsDirectoryW(bufW, CMDLINESIZE);
			break;

		case 1:
			GetSystemDirectoryW(bufW, CMDLINESIZE);
			break;

		default:
			GetTempPathW(CMDLINESIZE, bufW);
			break;
	}
	return JS_NewStringW(ctx, bufW);
}

static const JSCFunctionListEntry FSOFunctionList[] = {
	JS_CFUNC_DEF("BuildPath", 2, FsoBuildPath),
	JS_CFUNC_DEF("CopyFile", 2, FsoCopyFile),
	JS_CFUNC_DEF("CopyFolder", 2, FsoCopyFolder),
	JS_CFUNC_DEF("CreateFolder", 1, FsoCreateFolder),
		JS_PROP_INT32_DEF("CreateTextFile", SCRIPTMODULEVER, 0),
	JS_CFUNC_DEF("DeleteFile", 1, FsoDeleteFile),
	JS_CFUNC_DEF("DeleteFolder", 1, FsoDeleteFolder),
		JS_PROP_INT32_DEF("DriveExists", SCRIPTMODULEVER, 0),
	JS_CFUNC_DEF("FileExists", 1, FsoFileExists),
	JS_CFUNC_DEF("FolderExists", 1, FsoFolderExists),
	JS_CFUNC_DEF("GetAbsolutePathName", 1, FsoGetAbsolutePathName),
	JS_CFUNC_DEF("GetBaseName", 1, FsoGetBaseName),
		JS_PROP_INT32_DEF("GetDrive", SCRIPTMODULEVER, 0),
	JS_CFUNC_DEF("GetDriveName", 1, FsoGetDriveName),
	JS_CFUNC_DEF("GetExtensionName", 1, FsoGetExtensionName),
	JS_CFUNC_DEF("GetFile", 1, FsoGetEntry),
	JS_CFUNC_DEF("GetFileName", 1, FsoGetFileName),
	JS_CFUNC_DEF("GetFolder", 1, FsoGetEntry),
	JS_CFUNC_DEF("GetParentFolderName", 1, FsoGetParentFolderName),
	JS_CFUNC_DEF("GetSpecialFolder", 1, FsoGetSpecialFolder),
	JS_CGETSET_DEF("GetTempName", FsoGetTempName, NULL),
	JS_CFUNC_DEF("MoveFile", 2, FsoMoveFile),
	JS_CFUNC_DEF("MoveFolder", 2, FsoMoveFolder),
		JS_PROP_INT32_DEF("OpenAsTextStream", SCRIPTMODULEVER, 0),
		JS_PROP_INT32_DEF("OpenTextFile", SCRIPTMODULEVER, 0),
		JS_PROP_INT32_DEF("Drives", SCRIPTMODULEVER, 0),
};

JSValue New_FileSystemObject(JSContext *ctx)
{
	JSValue newobj = JS_NewObject(ctx);
	JS_SetPropertyFunctionList(ctx, newobj, FSOFunctionList, _countof(FSOFunctionList));
	JS_PreventExtensions(ctx, newobj);
	return newobj;
}
