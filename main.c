#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <uxtheme.h>
#include <locale.h>
#include <commctrl.h>
#include <richedit.h>
#include <tchar.h>
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <sys/types.h>
#include <math.h>
#include "xml.h"

#define LVS_EX_AUTOSIZECOLUMNS 0x10000000

#define WMU_UPDATE_GRID        WM_USER + 1
#define WMU_UPDATE_RESULTSET   WM_USER + 2
#define WMU_UPDATE_FILTER_SIZE WM_USER + 3
#define WMU_SET_HEADER_FILTERS WM_USER + 4
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 5
#define WMU_SET_CURRENT_CELL   WM_USER + 6
#define WMU_RESET_CACHE        WM_USER + 7
#define WMU_SET_FONT           WM_USER + 8
#define WMU_SET_THEME          WM_USER + 9
#define WMU_UPDATE_TEXT        WM_USER + 10
#define WMU_UPDATE_HIGHLIGHT   WM_USER + 11
#define WMU_SWITCH_TAB         WM_USER + 12
#define WMU_HIDE_COLUMN        WM_USER + 13
#define WMU_SHOW_COLUMNS       WM_USER + 14
#define WMU_SORT_COLUMN        WM_USER + 15
#define WMU_HOT_KEYS           WM_USER + 16  
#define WMU_HOT_CHARS          WM_USER + 17

#define IDC_MAIN               100
#define IDC_TREE               101
#define IDC_TAB                102
#define IDC_GRID               103
#define IDC_TEXT               104
#define IDC_STATUSBAR          105
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROWS          5001
#define IDM_COPY_COLUMN        5002
#define IDM_FILTER_ROW         5003
#define IDM_DARK_THEME         5004
#define IDM_COPY_TEXT          5010
#define IDM_SELECTALL          5011
#define IDM_FORMAT             5012
#define IDM_LOCATE             5013
#define IDM_COPY_XPATH         5020
#define IDM_HIDE_COLUMN        5021
#define IDM_SHOW_SAME          5022

#define SB_VERSION             0
#define SB_CODEPAGE            1
#define SB_RESERVED            2
#define SB_MODE                3
#define SB_ROW_COUNT           4
#define SB_CURRENT_CELL        5
#define SB_AUXILIARY           6

#define SPLITTER_WIDTH         5
#define MAX_LENGTH             4096
#define MAX_COLUMN_LENGTH      2000
#define APP_NAME               TEXT("xmltab")
#define APP_VERSION            TEXT("1.0.8")
#define LOADING                TEXT("Loading...")
#define WHITESPACE             " \t\r\n"

#define XML_TEXT               "#TEXT"
#define XML_COMMENT            "#COMMENT"
#define XML_CDATA              "#CDATA"

#define CP_UTF16LE             1200
#define CP_UTF16BE             1201

#define LCS_FINDFIRST          1
#define LCS_MATCHCASE          2
#define LCS_WHOLEWORDS         4
#define LCS_BACKWARDS          8

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = {0};

LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewTab(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewText(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

HTREEITEM addNode(HWND hTreeWnd, HTREEITEM hParentItem, xml_element* val);
void showNode(HWND hTreeWnd, xml_element* node);
void highlightText(HWND hWnd, TCHAR* text);
char* formatXML(const char* data);
HWND getMainWindow(HWND hWnd);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords);
int findString8(char* text, char* word, BOOL isMatchCase, BOOL isWholeWords);
BOOL hasString (const TCHAR* str, const TCHAR* sub, BOOL isCaseSensitive);
TCHAR* extractUrl(TCHAR* data);
int detectCodePage(const unsigned char *data);
void setClipboardText(const TCHAR* text);
BOOL isNumber(const TCHAR* val);
BOOL isEmpty (const char* s);
BOOL isUtf8(const char * string);
void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums);
HTREEITEM TreeView_AddItem (HWND hTreeWnd, TCHAR* caption, HTREEITEM parent, LPARAM lParam);
int TreeView_GetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* buf, int maxLen);
LPARAM TreeView_GetItemParam(HWND hTreeWnd, HTREEITEM hItem);
int TreeView_SetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* text);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	if (ul_reason_for_call == DLL_PROCESS_ATTACH && iniPath[0] == 0) {
		TCHAR path[MAX_PATH + 1] = {0};
		GetModuleFileName(hModule, path, MAX_PATH);
		TCHAR* dot = _tcsrchr(path, TEXT('.'));
		_tcsncpy(dot, TEXT(".ini"), 5);
		if (_taccess(path, 0) == 0)
			_tcscpy(iniPath, path);	
	}
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	TCHAR* detectString16 = getStoredString(TEXT("detect-string"), TEXT("MULTIMEDIA & ext=\"XML\""));
	char* detectString8 = utf16to8(detectString16);
	snprintf(DetectString, maxlen, detectString8);
	free(detectString16);
	free(detectString8);
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

int __stdcall ListSearchTextW(HWND hWnd, TCHAR* searchString, int searchParameter) {
	HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
	HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);

	BOOL isFindFirst = searchParameter & LCS_FINDFIRST;		
	BOOL isBackward = searchParameter & LCS_BACKWARDS;
	BOOL isMatchCase = searchParameter & LCS_MATCHCASE;
	BOOL isWholeWords = searchParameter & LCS_WHOLEWORDS;

	if (GetFocus() == hTreeWnd) {
		if (isBackward) {
			MessageBox(hWnd, TEXT("Backward search is unsupported"), NULL, MB_OK);
			return 0;
		}
		
		xml_element* xml = (xml_element*)GetProp(hWnd, TEXT("XML"));	
		HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);
		xml_element* from = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
		
		char* searchString8 = utf16to8(searchString);
		char* xml8 = (char*)GetProp(hWnd, TEXT("DATA"));
		int offset[] = {
			from->offset + strlen(from->key ? from->key : ""),
			from->offset + from->length
		};
		int pos[] = {
			offset[0] + findString8(xml8 + offset[0], searchString8, isMatchCase, isWholeWords),
			offset[1] + findString8(xml8 + offset[1], searchString8, isMatchCase, isWholeWords)
		};
		free(searchString8);
	
		if (pos[0] == offset[0] - 1)	
			return 0;
	
		xml_element* node = 0;	
		for (int i = 0; i < 2; i++) {
			node = xml;
			xml_element* prev = 0;
			while (node && prev != node) {
				prev = node;
	
				xml_element* next = node->next;
				while (next && (next->offset < pos[i])) {
					if(next->key)
						node = next;

					next = next->next;
				}
	
				xml_element* subnode = node->first_child;
				while (subnode && (subnode->offset < pos[i])) {
					if (subnode->key) 
						node = subnode;

					subnode = subnode->next;
				}
			}
	
			if (node != from)
				break;
		}
				
		showNode(hTreeWnd, node);
		TreeView_SelectItem(hTreeWnd, (HTREEITEM)(node ? node->userdata : 0));
	} else if (TabCtrl_GetCurSel(hTabWnd) == 1) { 
		HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
		DWORD len = _tcslen(searchString);
		DWORD spos = 0;
		SendMessage(hTextWnd, EM_GETSEL, 0, (LPARAM)&spos);
		int mode = 0;	
		FINDTEXTEXW ft = {{spos, -1}, searchString, {0, 0}};
		if (searchParameter & LCS_MATCHCASE)
			mode |= FR_MATCHCASE;
		if (searchParameter & LCS_WHOLEWORDS)
			mode |= FR_WHOLEWORD;
		if (!(searchParameter & LCS_BACKWARDS)) 
			mode |= FR_DOWN;
		else 
			ft.chrg.cpMin  = ft.chrg.cpMin > len ? ft.chrg.cpMin - len : ft.chrg.cpMin;
	
		int pos = SendMessage(hTextWnd, EM_FINDTEXTEXW, mode, (LPARAM)&ft);	
		if (pos != -1) 
			SendMessage(hTextWnd, EM_SETSEL, pos, pos + len);
		else	
			MessageBeep(0);
		SetFocus(hTextWnd);		
	} else {
		HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);	
		
		TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
		int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
		int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
		if (!resultset || rowCount == 0)
			return 0;
	
		if (isFindFirst) {
			*(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) = 0;
			*(int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")) = 0;	
			*(int*)GetProp(hWnd, TEXT("CURRENTROWNO")) = isBackward ? rowCount - 1 : 0;
		}	
			
		int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
		int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
		int *pStartPos = (int*)GetProp(hWnd, TEXT("SEARCHCELLPOS"));	
		rowNo = rowNo == -1 || rowNo >= rowCount ? 0 : rowNo;
		colNo = colNo == -1 || colNo >= colCount ? 0 : colNo;	
				
		int pos = -1;
		do {
			for (; (pos == -1) && colNo < colCount; colNo++) {
				pos = findString(cache[resultset[rowNo]][colNo] + *pStartPos, searchString, isMatchCase, isWholeWords);
				if (pos != -1) 
					pos += *pStartPos;				
				*pStartPos = pos == -1 ? 0 : pos + *pStartPos + _tcslen(searchString);
			}
			colNo = pos != -1 ? colNo - 1 : 0;
			rowNo += pos != -1 ? 0 : isBackward ? -1 : 1; 	
		} while ((pos == -1) && (isBackward ? rowNo > 0 : rowNo < rowCount));
		ListView_SetItemState(hGridWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
	
		TCHAR buf[256] = {0};
		if (pos != -1) {
			ListView_EnsureVisible(hGridWnd, rowNo, FALSE);
			ListView_SetItemState(hGridWnd, rowNo, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
			
			TCHAR* val = cache[resultset[rowNo]][colNo];
			int len = _tcslen(searchString);
			_sntprintf(buf, 255, TEXT("%ls%.*ls%ls"),
				pos > 0 ? TEXT("...") : TEXT(""), 
				len + pos + 10, val + pos,
				_tcslen(val + pos + len) > 10 ? TEXT("...") : TEXT(""));
			SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
		} else { 
			MessageBox(hWnd, searchString, TEXT("Not found:"), MB_OK);
		}
		SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)buf);	
		SetFocus(hGridWnd);
	}

	return 0;
}

int __stdcall ListSearchText(HWND hWnd, char* searchString, int searchParameter) {
	DWORD len = MultiByteToWideChar(CP_ACP, 0, searchString, -1, NULL, 0);
	TCHAR* searchString16 = (TCHAR*)calloc (len, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, searchString, -1, searchString16, len);
	int rc = ListSearchTextW(hWnd, searchString16, searchParameter);
	free(searchString16);
	return rc;
}

HWND APIENTRY ListLoadW (HWND hListerWnd, TCHAR* fileToLoad, int showFlags) {
	int maxFileSize = getStoredValue(TEXT("max-file-size"), 100000000);
	struct _stat st = {0};
	if (_tstat(fileToLoad, &st) != 0 || maxFileSize > 0 && st.st_size > maxFileSize || st.st_size < 4)
		return 0;

	char* data = calloc(st.st_size + 1, sizeof(char));
	FILE *f = _tfopen(fileToLoad, TEXT("rb"));
	fread(data, sizeof(char), st.st_size, f);
	fclose(f);
	
	int cp = detectCodePage(data);
	if (cp == CP_UTF16BE) {
		for (int i = 0; i < st.st_size/2; i++) {
			int c = data[2 * i];
			data[2 * i] = data[2 * i + 1];
			data[2 * i + 1] = c;
		}
	}
	
	if (cp == CP_UTF16LE || cp == CP_UTF16BE) {
		char* data8 = utf16to8((TCHAR*)data);
		free(data);
		data = data8;
	}
	
	if (cp == CP_ACP) {
		DWORD len = MultiByteToWideChar(CP_ACP, 0, data, -1, NULL, 0);
		TCHAR* data16 = (TCHAR*)calloc (len, sizeof (TCHAR));
		MultiByteToWideChar(CP_ACP, 0, data, -1, data16, len);
		free(data);
		char* data8 = utf16to8((TCHAR*)data16);
		data = data8;		
		free(data16);		
	}
		
	struct xml_element *xml = xml_parse(data);
	
	// If doc doesn't have a root, add it
	int nCount = 0;
	if (xml) {
		xml_element *node = xml->first_child;
		while (node) {
			nCount += node->key && strchr("?! ", node->key[0]) == 0;
			node = node->next;
		}
	}
		
	if (nCount > 1) {
		char* data2 = calloc(st.st_size + 64, sizeof(char));
		sprintf(data2, "<root$>%s</root$>", data);
		free(data);
		xml_free(xml);
		
		data = data2;
		xml = xml_parse(data);
	}
	
	if (!xml) {
		free(data);
		return 0;
	}

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	LoadLibrary(TEXT("msftedit.dll"));
	
	setlocale(LC_CTYPE, "");

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindowEx(WS_EX_CONTROLPARENT, WC_STATIC, APP_NAME, WS_CHILD | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("FILTERROW"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("XML"), xml);
	SetProp(hMainWnd, TEXT("DATA"), data);	
	SetProp(hMainWnd, TEXT("CACHE"), 0);	
	SetProp(hMainWnd, TEXT("RESULTSET"), 0);
	SetProp(hMainWnd, TEXT("XMLNODES"), 0);	
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("CURRENTROWNO"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCOLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SEARCHCELLPOS"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERPOSITION"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ISFORMAT"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));	
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERALIGN"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("MAXHIGHLIGHTLENGTH"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SAMEFILTER"), calloc(255, sizeof(TCHAR)));	
	
	SetProp(hMainWnd, TEXT("DARKTHEME"), calloc(1, sizeof(int)));			
	SetProp(hMainWnd, TEXT("TEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR2"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERBACKCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCELLCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SELECTIONTEXTCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SELECTIONBACKCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SPLITTERCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("XMLTEXTCOLOR"), calloc(1, sizeof(int)));		
	SetProp(hMainWnd, TEXT("XMLTAGCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("XMLSTRINGCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("XMLVALUECOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("XMLCDATACOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("XMLCOMMENTCOLOR"), calloc(1, sizeof(int)));

	*(int*)GetProp(hMainWnd, TEXT("SPLITTERPOSITION")) = getStoredValue(TEXT("splitter-position"), 200);
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);
	*(int*)GetProp(hMainWnd, TEXT("ISFORMAT")) = getStoredValue(TEXT("format"), 1);
	*(int*)GetProp(hMainWnd, TEXT("MAXHIGHLIGHTLENGTH")) = getStoredValue(TEXT("max-highlight-length"), 64000);	
	*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) = getStoredValue(TEXT("filter-row"), 1);
	*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) = getStoredValue(TEXT("dark-theme"), 0);
	*(int*)GetProp(hMainWnd, TEXT("FILTERALIGN")) = getStoredValue(TEXT("filter-align"), 0);	

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE |  (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hMainWnd);
	float z = GetDeviceCaps(hDC, LOGPIXELSX) / 96.0; // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hMainWnd, hDC);		
	int sizes[7] = {35 * z, 95 * z, 150 * z, 200 * z, 400 * z, 500 * z, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&sizes);

	HWND hTreeWnd = CreateWindow(WC_TREEVIEW, NULL, WS_VISIBLE | WS_CHILD | TVS_HASBUTTONS | TVS_HASLINES | TVS_SHOWSELALWAYS | WS_TABSTOP,
		0, 0, 100, 100, hMainWnd, (HMENU)IDC_TREE, GetModuleHandle(0), NULL);
	SetProp(hTreeWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTreeWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	HWND hTabWnd = CreateWindow(WC_TABCONTROL, NULL, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | WS_TABSTOP, 100, 100, 100, 100,
		hMainWnd, (HMENU)IDC_TAB, GetModuleHandle(0), NULL);
	SetProp(hTabWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTabWnd, GWLP_WNDPROC, (LONG_PTR)cbNewTab));	

	TCITEM tci;
	tci.mask = TCIF_TEXT | TCIF_IMAGE;
	tci.iImage = -1;
	tci.pszText = TEXT("Grid");
	tci.cchTextMax = 5;
	TabCtrl_InsertItem(hTabWnd, 0, &tci);

	tci.pszText = TEXT("Text");
	tci.cchTextMax = 5;
	TabCtrl_InsertItem(hTabWnd, 1, &tci);

	int tabNo = getStoredValue(TEXT("tab-no"), 0);
	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, (tabNo == 0 ? WS_VISIBLE : 0) | WS_CHILD  | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hTabWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
	
	int noLines = getStoredValue(TEXT("disable-grid-lines"), 0);
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | (noLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP | LVS_EX_HEADERDRAGDROP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));

	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)cbNewHeader));

	HWND hTextWnd = CreateWindowEx(0, TEXT("RICHEDIT50W"), NULL, (tabNo == 1 ? WS_VISIBLE : 0) | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN | WS_VSCROLL | WS_HSCROLL | WS_TABSTOP | ES_NOHIDESEL | ES_READONLY,
		0, 0, 100, 100, hTabWnd, (HMENU)IDC_TEXT, GetModuleHandle(0),  NULL);
	SetProp(hTextWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTextWnd, GWLP_WNDPROC, (LONG_PTR)cbNewText));
	TabCtrl_SetCurSel(hTabWnd, tabNo);

	HMENU hTreeMenu = CreatePopupMenu();
	AppendMenu(hTreeMenu, MF_STRING, IDM_COPY_XPATH, TEXT("Copy XPath"));
	AppendMenu(hTreeMenu, MF_STRING, IDM_SHOW_SAME, TEXT("Show the same siblings"));	
	SetProp(hMainWnd, TEXT("TREEMENU"), hTreeMenu);
	
	HMENU hGridMenu = CreatePopupMenu();
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_ROWS, TEXT("Copy row(s)"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_COLUMN, TEXT("Copy column"));	
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_XPATH, TEXT("Copy XPath"));		
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);
	AppendMenu(hGridMenu, MF_STRING, IDM_HIDE_COLUMN, TEXT("Hide column"));	
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_FILTER_ROW, TEXT("Filters"));	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));	
	SetProp(hMainWnd, TEXT("GRIDMENU"), hGridMenu);

	HMENU hTextMenu = CreatePopupMenu();
	AppendMenu(hTextMenu, MF_STRING, IDM_COPY_TEXT, TEXT("Copy"));
	AppendMenu(hTextMenu, MF_STRING, IDM_SELECTALL, TEXT("Select all"));
	AppendMenu(hTextMenu, MF_STRING, 0, NULL);
	AppendMenu(hTextMenu, MF_STRING | (*(int*)GetProp(hMainWnd, TEXT("ISFORMAT")) != 0 ? MF_CHECKED : 0), IDM_FORMAT, TEXT("Format"));
	AppendMenu(hTextMenu, MF_STRING, IDM_LOCATE, TEXT("Locate"));		
	AppendMenu(hTextMenu, MF_STRING, 0, NULL);	
	AppendMenu(hTextMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));						
	SetProp(hMainWnd, TEXT("TEXTMENU"), hTextMenu);
	
	TCHAR buf[255];
	_sntprintf(buf, 32, TEXT(" %ls"), APP_VERSION);	
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)buf);	
	SendMessage(hStatusWnd, SB_SETTEXT, SB_CODEPAGE, 
		(LPARAM)(cp == CP_UTF8 ? TEXT("    UTF-8") : cp == CP_UTF16LE ? TEXT(" UTF-16LE") : cp == CP_UTF16BE ? TEXT(" UTF-16BE") : TEXT("     ANSI")));	
	
	xml_element* node = xml->last_child;
	while (node) {
		HTREEITEM hItem = addNode(hTreeWnd, TVI_ROOT, node);
		node->userdata = (void*)hItem;	
		TreeView_Expand(hTreeWnd, hItem, TVE_EXPAND);		
		node = node->prev;
	}

	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);	
	SendMessage(hMainWnd, WMU_SET_THEME, 0, 0);	
	HTREEITEM hItem = TreeView_GetNextItem(hTreeWnd, TVI_ROOT, TVGN_CHILD);
	if (getStoredValue(TEXT("open-first-element"), 1)) {
		xml_element* node = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
		while(node && (!node->key || node->key[0] == '?' || node->key[0] == '!'))
			node = node->next;
		
		if (node)
			hItem = (HTREEITEM)node->userdata;
	}
	TreeView_Select(hTreeWnd, hItem, TVGN_CARET);
	ShowWindow(hMainWnd, SW_SHOW);
	SetFocus(hTreeWnd);

	return hMainWnd;
}

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* fileToLoadW = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, fileToLoadW, size);
	HWND hWnd = ListLoadW(hListerWnd, fileToLoadW, showFlags);
	free(fileToLoadW);
	return hWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("splitter-position"), *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("format"), *(int*)GetProp(hWnd, TEXT("ISFORMAT")));	
	setStoredValue(TEXT("tab-no"), TabCtrl_GetCurSel(GetDlgItem(hWnd, IDC_TAB)));
	setStoredValue(TEXT("filter-row"), *(int*)GetProp(hWnd, TEXT("FILTERROW")));	
	setStoredValue(TEXT("dark-theme"), *(int*)GetProp(hWnd, TEXT("DARKTHEME")));

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	xml_free((xml_element*)GetProp(hWnd, TEXT("XML")));
	free((int*)GetProp(hWnd, TEXT("FILTERROW")));
	free((int*)GetProp(hWnd, TEXT("DARKTHEME")));	
	free((char*)GetProp(hWnd, TEXT("DATA")));
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("CURRENTROWNO")));				
	free((int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));
	free((int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")));
	free((int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	free((int*)GetProp(hWnd, TEXT("ISFORMAT")));	
	free((int*)GetProp(hWnd, TEXT("MAXHIGHLIGHTLENGTH")));
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
	free((int*)GetProp(hWnd, TEXT("FILTERALIGN")));		
	free((TCHAR*)GetProp(hWnd, TEXT("SAMEFILTER")));	

	free((int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR2")));	
	free((int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")));
	free((int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")));
	free((int*)GetProp(hWnd, TEXT("XMLTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("XMLTAGCOLOR")));
	free((int*)GetProp(hWnd, TEXT("XMLSTRINGCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("XMLVALUECOLOR")));
	free((int*)GetProp(hWnd, TEXT("XMLCDATACOLOR")));
	free((int*)GetProp(hWnd, TEXT("XMLCOMMENTCOLOR")));	

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));	
	DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
	DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("TREEMENU")));	
	DestroyMenu(GetProp(hWnd, TEXT("GRIDMENU")));
	DestroyMenu(GetProp(hWnd, TEXT("TEXTMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("FILTERROW"));	
	RemoveProp(hWnd, TEXT("DARKTHEME"));	
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("RESULTSET"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("CURRENTROWNO"));	
	RemoveProp(hWnd, TEXT("CURRENTCOLNO"));
	RemoveProp(hWnd, TEXT("SEARCHCELLPOS"));		
	RemoveProp(hWnd, TEXT("XML"));
	RemoveProp(hWnd, TEXT("DATA"));	
	RemoveProp(hWnd, TEXT("SPLITTERPOSITION"));
	RemoveProp(hWnd, TEXT("ISFORMAT"));
	RemoveProp(hWnd, TEXT("MAXHIGHLIGHTLENGTH"));
	RemoveProp(hWnd, TEXT("FILTERALIGN"));
	RemoveProp(hWnd, TEXT("LASTFOCUS"));		
	RemoveProp(hWnd, TEXT("SAMEFILTER"));	
	
	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));	
	RemoveProp(hWnd, TEXT("FONTSIZE"));	
	RemoveProp(hWnd, TEXT("TEXTCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR2"));	
	RemoveProp(hWnd, TEXT("FILTERTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("FILTERBACKCOLOR"));
	RemoveProp(hWnd, TEXT("CURRENTCELLCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONBACKCOLOR"));
	RemoveProp(hWnd, TEXT("SPLITTERCOLOR"));
	RemoveProp(hWnd, TEXT("XMLTEXTCOLOR"));	
	RemoveProp(hWnd, TEXT("XMLTAGCOLOR"));		
	RemoveProp(hWnd, TEXT("XMLSTRINGCOLOR"));			
	RemoveProp(hWnd, TEXT("XMLVALUECOLOR"));		
	RemoveProp(hWnd, TEXT("XMLCDATACOLOR"));		
	RemoveProp(hWnd, TEXT("XMLCOMMENTCOLOR"));			
	RemoveProp(hWnd, TEXT("BACKBRUSH"));
	RemoveProp(hWnd, TEXT("FILTERBACKBRUSH"));		
	RemoveProp(hWnd, TEXT("SPLITTERBRUSH"));
	RemoveProp(hWnd, TEXT("TREEMENU"));	
	RemoveProp(hWnd, TEXT("GRIDMENU"));
	RemoveProp(hWnd, TEXT("TEXTMENU"));

	DestroyWindow(hWnd);
}

LRESULT CALLBACK cbNewMain(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			int splitterW = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			GetClientRect(hWnd, &rc);
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			SetWindowPos(hTreeWnd, 0, 0, 0, splitterW, rc.bottom - statusH, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(hTabWnd, 0, splitterW + SPLITTER_WIDTH, 0, rc.right - splitterW - SPLITTER_WIDTH, rc.bottom - statusH, SWP_NOZORDER);

			RECT rc2;
			GetClientRect(hTabWnd, &rc);
			TabCtrl_GetItemRect(hTabWnd, 0, &rc2);
			SetWindowPos(hTextWnd, 0, 2, rc2.bottom + 3, rc.right - rc.left - SPLITTER_WIDTH, rc.bottom - rc2.bottom - 7, SWP_NOZORDER);
			SetWindowPos(hGridWnd, 0, 2, rc2.bottom + 3, rc.right - rc.left - SPLITTER_WIDTH, rc.bottom - rc2.bottom - 7, SWP_NOZORDER);
		}
		break;

		case WM_PAINT: {
			PAINTSTRUCT ps = {0};
			HDC hDC = BeginPaint(hWnd, &ps);

			RECT rc;
			GetClientRect(hWnd, &rc);
			rc.left = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			rc.right = rc.left + SPLITTER_WIDTH;
			FillRect(hDC, &rc, (HBRUSH)GetProp(hWnd, TEXT("SPLITTERBRUSH")));
			EndPaint(hWnd, &ps);

			return 0;
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;
		
		case WM_SETCURSOR: {
			SetCursor(LoadCursor(0, GetProp(hWnd, TEXT("ISMOUSEHOVER")) ? IDC_SIZEWE : IDC_ARROW));
			return TRUE;
		}
		break;	
		
		case WM_SETFOCUS: {
			SetFocus(GetProp(hWnd, TEXT("LASTFOCUS")));
		}
		break;			

		case WM_LBUTTONDOWN: {
			int x = GET_X_LPARAM(lParam);
			int pos = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			if (x >= pos && x <= pos + SPLITTER_WIDTH) {
				SetProp(hWnd, TEXT("ISMOUSEDOWN"), (HANDLE)1);
				SetCapture(hWnd);
			}
			return 0;
		}
		break;

		case WM_LBUTTONUP: {
			ReleaseCapture();
			RemoveProp(hWnd, TEXT("ISMOUSEDOWN"));
		}
		break;

		case WM_MOUSEMOVE: {
			DWORD x = GET_X_LPARAM(lParam);
			int* pPos = (int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			
			if (!GetProp(hWnd, TEXT("ISMOUSEHOVER")) && *pPos <= x && x <= *pPos + SPLITTER_WIDTH) {
				TRACKMOUSEEVENT tme = {sizeof(TRACKMOUSEEVENT), TME_LEAVE, hWnd, 0};
				TrackMouseEvent(&tme);	
				SetProp(hWnd, TEXT("ISMOUSEHOVER"), (HANDLE)1);
			}
			
			if (GetProp(hWnd, TEXT("ISMOUSEDOWN")) && x > 0 && x < 32000) {
				*pPos = x;
				SendMessage(hWnd, WM_SIZE, 0, 0);
			}
		}
		break;
		
		case WM_MOUSELEAVE: {
			SetProp(hWnd, TEXT("ISMOUSEHOVER"), 0);
		}
		break;

		case WM_MOUSEWHEEL: {
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1: -1, 0);
				return 1;
			}
		}
		break;
		
		case WM_KEYDOWN: {
			if (SendMessage(hWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;
		
		case WM_CONTEXTMENU: {
			POINT p = {LOWORD(lParam), HIWORD(lParam)};
			int id = GetDlgCtrlID(WindowFromPoint(p));
			if (id == IDC_TEXT) {
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
				ScreenToClient(hTextWnd, &p);
				int pos = SendMessage(hTextWnd, EM_CHARFROMPOS, 0, (LPARAM)&p);
				int start, end;
				SendMessage(hTextWnd, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
				if (start == end || pos < start || pos > end)
					SendMessage(hTextWnd, EM_SETSEL, pos, pos);
				
				ClientToScreen(hTextWnd, &p);
				TrackPopupMenu(GetProp(hWnd, TEXT("TEXTMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}
			
			// update selected item on right click
			if (id == IDC_TREE) {
				HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);				
				POINT p2 = {0};
				GetCursorPos(&p2);
				ScreenToClient(hTreeWnd, &p2);
				TVHITTESTINFO thi = {p2, TVHT_ONITEM};
				HTREEITEM hItem = TreeView_HitTest(hTreeWnd, &thi);
				TreeView_SelectItem(hTreeWnd, hItem);
				if (!hItem)
					return 0;
				
				TCHAR nodeName[256] = {0};
				TreeView_GetItemText(hTreeWnd, hItem, nodeName, 255);
	
				HMENU hTreeMenu = GetProp(hWnd, TEXT("TREEMENU")); 
				Menu_SetItemState(hTreeMenu, IDM_SHOW_SAME, _istalpha(nodeName[0]) ? MFS_ENABLED : MFS_DISABLED);		
				TrackPopupMenu(hTreeMenu, TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROWS || cmd == IDM_COPY_COLUMN) {
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				
				int colCount = Header_GetItemCount(hHeader);
				int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
				int selCount = ListView_GetSelectedCount(hGridWnd);

				if (rowNo == -1 ||
					rowNo >= rowCount ||
					colCount == 0 ||
					cmd == IDM_COPY_CELL && colNo == -1 || 
					cmd == IDM_COPY_CELL && colNo >= colCount || 
					cmd == IDM_COPY_COLUMN && colNo == -1 || 
					cmd == IDM_COPY_COLUMN && colNo >= colCount || 					
					cmd == IDM_COPY_ROWS && selCount == 0) {
					setClipboardText(TEXT(""));
					return 0;
				}
						
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));				
				if (!resultset)
					return 0;

				TCHAR* delimiter = getStoredString(TEXT("column-delimiter"), TEXT("\t"));

				int len = 0;
				if (cmd == IDM_COPY_CELL) 
					len = _tcslen(cache[resultset[rowNo]][colNo]);
				
				if (cmd == IDM_COPY_ROWS) {
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hGridWnd, colNo)) 
								len += _tcslen(cache[resultset[rowNo]][colNo]) + 1; /* column delimiter */
						}
													
						len++; /* \n */		
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
				}

				if (cmd == IDM_COPY_COLUMN) {
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						len += _tcslen(cache[resultset[rowNo]][colNo]) + 1 /* \n */;
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
					} 
				}	

				TCHAR* buf = calloc(len + 1, sizeof(TCHAR));
				if (cmd == IDM_COPY_CELL)
					_tcscat(buf, cache[resultset[rowNo]][colNo]);
				
				if (cmd == IDM_COPY_ROWS) {
					int pos = 0;
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);

					int* colOrder = calloc(colCount, sizeof(int));
					Header_GetOrderArray(hHeader, colCount, colOrder);

					while (rowNo != -1) {
						for (int idx = 0; idx < colCount; idx++) {
							int colNo = colOrder[idx];
							if (ListView_GetColumnWidth(hGridWnd, colNo)) {
								int len = _tcslen(cache[resultset[rowNo]][colNo]);
								_tcsncpy(buf + pos, cache[resultset[rowNo]][colNo], len);
								buf[pos + len] = delimiter[0];
								pos += len + 1;
							}
						}

						buf[pos - (pos > 0)] = TEXT('\n');
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
					buf[pos - 1] = 0; // remove last \n

					free(colOrder);
				}

				if (cmd == IDM_COPY_COLUMN) {
					int pos = 0;
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						int len = _tcslen(cache[resultset[rowNo]][colNo]);
						_tcsncpy(buf + pos, cache[resultset[rowNo]][colNo], len);
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
						if (rowNo != -1 && rowNo < rowCount)
							buf[pos + len] = TEXT('\n');
						pos += len + 1;								
					} 
				}
									
				setClipboardText(buf);
				free(buf);
				free(delimiter);
			}

			if (cmd == IDM_COPY_TEXT || cmd == IDM_SELECTALL) {
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
				SendMessage(hTextWnd, cmd == IDM_COPY_TEXT ? WM_COPY : EM_SETSEL, 0, -1);
			}
			
			if (cmd == IDM_FORMAT) {
				HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("TEXTMENU"));
				int* pIsFormat = (int*)GetProp(hWnd, TEXT("ISFORMAT"));
				*pIsFormat = (*pIsFormat + 1) % 2;
				
				MENUITEMINFO mii = {0};
				mii.cbSize = sizeof(MENUITEMINFO);
				mii.fMask = MIIM_STATE;
				mii.fState = *pIsFormat ? MFS_CHECKED : 0;
				SetMenuItemInfo(hMenu, IDM_FORMAT, FALSE, &mii);				
				
				SendMessage(hWnd, WMU_UPDATE_TEXT, 0, 0);
			}
			
			if (cmd == IDM_LOCATE) {
				HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);				
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
				int pos16; 
				SendMessage(hTextWnd, EM_GETSEL, (WPARAM)&pos16, 0);

				GETTEXTLENGTHEX gtl = {GTL_NUMBYTES, 0};
				int len = SendMessage(hTextWnd, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 1200);
				TCHAR* xml16 = calloc(len + sizeof(TCHAR), sizeof(char));
				GETTEXTEX gt = {0};
				gt.cb = len + sizeof(TCHAR);
				gt.flags = 0;
				gt.codepage = 1200;
				SendMessage(hTextWnd, EM_GETTEXTEX, (WPARAM)&gt, (LPARAM)xml16);
																							
				char* xml8 = utf16to8(xml16);			
				char* xml8_ = utf16to8(xml16 + pos16);
				int pos8 = strlen(xml8) - strlen(xml8_);
				free(xml8_);
								
				xml_element* xml = xml_parse(xml8);
				xml_element* node = xml->first_child;
				HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);				
				
				while (node) {
					if (node->next && node->next->offset < pos8)  {
						node = node->next;
						if (node->key || node->value && !isEmpty(node->value)) 
							hItem = TreeView_GetNextSibling(hTreeWnd, hItem) ? : hItem;
						
						continue;
					} 
					
					if (node->first_child && node->first_child->offset < pos8) {
						node = node->first_child;
						TreeView_Expand(hTreeWnd, hItem, TVE_EXPAND);
						while (!(node->key || node->value && !isEmpty(node->value)))
							node = node->next;
						hItem = TreeView_GetChild(hTreeWnd, hItem) ? : hItem;						
						
						continue;
					} 
					
					break;
				}
																
				xml_free(xml);
				free(xml8);
				free(xml16);
				
				if (hItem)
					TreeView_SelectItem(hTreeWnd, hItem);
				else
					MessageBeep(0);	
			}
			
			if (cmd == IDM_HIDE_COLUMN) {
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				SendMessage(hWnd, WMU_HIDE_COLUMN, colNo, 0);
			}
						
			if (cmd == IDM_SHOW_SAME) {				
				HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
				HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);
				if (hItem) {
					xml_element* node = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
					if (node->key && strlen(node->key)) {
						TCHAR* nodeName = utf8to16(node->key);
						_tcsncpy((TCHAR*)GetProp(hWnd, TEXT("SAMEFILTER")), nodeName, 255);
						free(nodeName);
					}
					TreeView_SelectItem(hTreeWnd, TreeView_GetParent(hTreeWnd, hItem));
				}
			}
			
			if (cmd == IDM_FILTER_ROW || cmd == IDM_DARK_THEME) {
				HMENU hGridMenu = (HMENU)GetProp(hWnd, TEXT("GRIDMENU"));
				HMENU hTextMenu = (HMENU)GetProp(hWnd, TEXT("TEXTMENU"));				
				int* pOpt = (int*)GetProp(hWnd, cmd == IDM_FILTER_ROW ? TEXT("FILTERROW") : TEXT("DARKTHEME"));
				*pOpt = (*pOpt + 1) % 2;
				Menu_SetItemState(hGridMenu, cmd, *pOpt ? MFS_CHECKED : 0);
				Menu_SetItemState(hTextMenu, cmd, *pOpt ? MFS_CHECKED : 0);				
				
				UINT msg = cmd == IDM_FILTER_ROW ? WMU_SET_HEADER_FILTERS : WMU_SET_THEME;
				SendMessage(hWnd, msg, 0, 0);				
			}
			
			if (cmd == IDM_COPY_XPATH) {
				HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);			
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);				

				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				xml_element* node = 0;
				TCHAR* attr16 = 0;
				if (GetFocus() == hGridWnd) {
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
					xml_element** xmlnodes = (xml_element**)GetProp(hWnd, TEXT("XMLNODES"));
					
					node = xmlnodes[resultset[rowNo]];
					if (!node) 
						attr16 = cache[rowNo][0];
				}
				
				if (!node) {
					HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);
					node = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
				}
				
				if (node) {
					char* xpath8 = xml_path(node);
					TCHAR* xpath16 = utf8to16(xpath8);
					int len = _tcslen(xpath16) + (attr16 ? _tcslen(attr16) : 0) + 10;
					TCHAR* buf16 = calloc(len + 1, sizeof(TCHAR));
					if (attr16)
						_sntprintf(buf16, len, TEXT("%ls/%ls"), xpath16, attr16);
					else
						_sntprintf(buf16, len, TEXT("%ls"), xpath16);
						
					setClipboardText(buf16);
					
					free(buf16);	
					free(xpath8);
					free(xpath16);
				} else {
					MessageBeep(0);
				}
				
			}			
		}
		break;

		case WM_NOTIFY : {
			NMHDR* pHdr = (LPNMHDR)lParam;			
			if (pHdr->idFrom == IDC_TAB && pHdr->code == TCN_SELCHANGE) {
				HWND hTabWnd = pHdr->hwndFrom;
				BOOL isText = TabCtrl_GetCurSel(hTabWnd) == 1;
				ShowWindow(GetDlgItem(hTabWnd, IDC_GRID), isText ? SW_HIDE : SW_SHOW);
				ShowWindow(GetDlgItem(hTabWnd, IDC_TEXT), isText ? SW_SHOW : SW_HIDE);
			}

			if (pHdr->idFrom == IDC_TREE && pHdr->code == NM_CLICK) {
				HWND hTreeWnd = pHdr->hwndFrom;
				POINT p = {0};
				GetCursorPos(&p);
				ScreenToClient(hTreeWnd, &p);
				TVHITTESTINFO thi = {p, TVHT_ONITEM};
				if (TreeView_HitTest(hTreeWnd, &thi) == TreeView_GetSelection(hTreeWnd)) {
					SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
					SendMessage(hWnd, WMU_UPDATE_TEXT, 0, 0);
				}
			}		
			
			if (pHdr->idFrom == IDC_TREE && pHdr->code == TVN_SELCHANGED) {
				SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
				SendMessage(hWnd, WMU_UPDATE_TEXT, 0, 0);
			}
				
			if (pHdr->idFrom == IDC_TREE && pHdr->code == TVN_ITEMEXPANDING) {
				HWND hTreeWnd = pHdr->hwndFrom;
				NMTREEVIEW* tv = (NMTREEVIEW*) lParam;
				HTREEITEM hItem = tv->itemNew.hItem;				
				HTREEITEM hChild = TreeView_GetChild(hTreeWnd, hItem);

				if (hChild) {
					TCHAR nodeName[256];
					TreeView_GetItemText(hTreeWnd, hChild, nodeName, 255);
					if (_tcscmp(nodeName, LOADING) == 0) {
						TreeView_DeleteItem(hTreeWnd, hChild);
						
						xml_element* node = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
						xml_element* subnode = node != NULL ? node->last_child : 0; 
						while (subnode) {
							subnode->userdata = (void*)addNode(hTreeWnd, hItem, subnode);
							subnode = subnode->prev;
						}
					}
				}
			}	

			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_GETDISPINFO) {
				LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
				LV_ITEM* pItem = &(pDispInfo)->item;

				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				if(resultset && pItem->mask & LVIF_TEXT) {
					int rowNo = resultset[pItem->iItem];
					pItem->pszText = cache[rowNo][pItem->iSubItem];
				}
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_COLUMNCLICK) {
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				return SendMessage(hWnd, HIWORD(GetKeyState(VK_CONTROL)) ? WMU_HIDE_COLUMN : WMU_SORT_COLUMN, lv->iSubItem, 0);												
			}

			if (pHdr->idFrom == IDC_GRID && (pHdr->code == (DWORD)NM_CLICK || pHdr->code == (DWORD)NM_RCLICK)) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				SendMessage(hWnd, WMU_SET_CURRENT_CELL, ia->iItem, ia->iSubItem);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {
				POINT p;
				GetCursorPos(&p);
				TrackPopupMenu(GetProp(hWnd, TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				if (lv->uOldState != lv->uNewState && (lv->uNewState & LVIS_SELECTED))				
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, lv->iItem, *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));	
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_CLICK && HIWORD(GetKeyState(VK_MENU))) {	
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				
				TCHAR* url = extractUrl(cache[resultset[ia->iItem]][ia->iSubItem]);
				ShellExecute(0, TEXT("open"), url, 0, 0 , SW_SHOW);
				free(url);
			}			

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43) { // C
					BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
					BOOL isShift = HIWORD(GetKeyState(VK_SHIFT)); 
					BOOL isCopyColumn = getStoredValue(TEXT("copy-column"), 0) && ListView_GetSelectedCount(pHdr->hwndFrom) > 1;
					if (!isCtrl && !isShift)
						return FALSE;
						
					int action = !isShift && !isCopyColumn ? IDM_COPY_CELL : isCtrl || isCopyColumn ? IDM_COPY_COLUMN : IDM_COPY_ROWS;
					SendMessage(hWnd, WM_COMMAND, action, 0);
					return TRUE;
				}

				if (kd->wVKey == 0x41 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + A
					HWND hGridWnd = pHdr->hwndFrom;
					SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));					
					ListView_SetItemState(hGridWnd, -1, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED);
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
					SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
					InvalidateRect(hGridWnd, NULL, TRUE);
				}
				
				if (kd->wVKey == 0x20 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + Space				
					SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);					
					return TRUE;
				}				
				
				if (kd->wVKey == VK_LEFT || kd->wVKey == VK_RIGHT) {
					HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
					HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
					HWND hHeader = ListView_GetHeader(hGridWnd);

					int colCount = Header_GetItemCount(ListView_GetHeader(pHdr->hwndFrom));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));

					int* colOrder = calloc(colCount, sizeof(int));
					Header_GetOrderArray(hHeader, colCount, colOrder);
	
					int dir = kd->wVKey == VK_RIGHT ? 1 : -1;
					int idx = 0;
					for (idx; colOrder[idx] != colNo; idx++);
					do {
						idx = (colCount + idx + dir) % colCount;
					} while (ListView_GetColumnWidth(hGridWnd, colOrder[idx]) == 0);

					colNo = colOrder[idx];
					free(colOrder);

					SendMessage(hWnd, WMU_SET_CURRENT_CELL, *(int*)GetProp(hWnd, TEXT("CURRENTROWNO")), colNo);
					return TRUE;
				}
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_DBLCLK) {
				NMITEMACTIVATE* ia = (NMITEMACTIVATE*) lParam;
				if (ia->iItem == -1)	
					return 0;
					
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				xml_element** xmlnodes = (xml_element**)GetProp(hWnd, TEXT("XMLNODES"));


				HWND hGridWnd = pHdr->hwndFrom;
				HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
				
				int rowNo = resultset[ia->iItem];
				xml_element* node = xmlnodes[rowNo];
				HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);
				TreeView_Expand(hTreeWnd, hItem, TVE_EXPAND);
				hItem = TreeView_GetChild(hTreeWnd, hItem);
				if (node) {
					while (hItem && node != (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem)) 
						hItem = TreeView_GetNextSibling(hTreeWnd, hItem); 
				}	

				if (hItem)
					TreeView_SelectItem(hTreeWnd, hItem);
			}			

			if ((pHdr->code == HDN_ITEMCHANGED || pHdr->code == HDN_ENDDRAG) && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(GetDlgItem(hWnd, IDC_TAB), IDC_GRID)))
				PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
				
			if (pHdr->code == (UINT)NM_SETFOCUS) 
				SetProp(hWnd, TEXT("LASTFOCUS"), pHdr->hwndFrom);
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (UINT)NM_CUSTOMDRAW) {
				int result = CDRF_DODEFAULT;
				
				NMLVCUSTOMDRAW* pCustomDraw = (LPNMLVCUSTOMDRAW)lParam;				
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT) 
					result = CDRF_NOTIFYITEMDRAW;
	
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					if (ListView_GetItemState(pHdr->hwndFrom, pCustomDraw->nmcd.dwItemSpec, LVIS_SELECTED)) {
						pCustomDraw->nmcd.uItemState &= ~CDIS_SELECTED;
						result = CDRF_NOTIFYSUBITEMDRAW;
					} else {
						pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, pCustomDraw->nmcd.dwItemSpec % 2 == 0 ? TEXT("BACKCOLOR") : TEXT("BACKCOLOR2"));
					}				
				}
				
				if (pCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) {
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
					BOOL isCurrCell = (pCustomDraw->nmcd.dwItemSpec == (DWORD)rowNo) && (pCustomDraw->iSubItem == colNo);
					pCustomDraw->clrText = *(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
					pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, isCurrCell ? TEXT("CURRENTCELLCOLOR") : TEXT("SELECTIONBACKCOLOR"));
				}
	
				return result;
			}				
		}
		break;
		
		// wParam = colNo
		case WMU_HIDE_COLUMN: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);		
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colNo = (int)wParam;

			HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
			SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hGridWnd, colNo));				
			ListView_SetColumnWidth(hGridWnd, colNo, 0); 
			InvalidateRect(hHeader, NULL, TRUE);			
		}
		break;
		
		case WMU_SHOW_COLUMNS: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			for (int colNo = 0; colNo < colCount; colNo++) {
				if (ListView_GetColumnWidth(hGridWnd, colNo) == 0) {
					HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
					ListView_SetColumnWidth(hGridWnd, colNo, (int)GetWindowLongPtr(hEdit, GWLP_USERDATA));
				}
			}

			InvalidateRect(hGridWnd, NULL, TRUE);		
		}
		break;

		// wParam = colNo
		case WMU_SORT_COLUMN: {
			int colNo = (int)wParam + 1;
			if (colNo <= 0)
				return FALSE;

			int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
			int orderBy = *pOrderBy;
			*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
			SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);	
		}	
		break;

		case WMU_UPDATE_GRID: {
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			int filterAlign = *(int*)GetProp(hWnd, TEXT("FILTERALIGN"));
			BOOL isShowContent = getStoredValue(TEXT("show-content"), 0);				
						
			HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);			
			xml_element* node = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
			if (!node)
				return 0;
				
			TCHAR* sameFilter16 = (TCHAR*)GetProp(hWnd, TEXT("SAMEFILTER"));
			char* sameFilter8 = utf16to8(sameFilter16);
			BOOL isSameFilter = strlen(sameFilter8) > 0;
				
			HWND hHeader = ListView_GetHeader(hGridWnd);
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			SendMessage(hWnd, WMU_SET_CURRENT_CELL, 0, 0);
			*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;

			int colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++)
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hGridWnd, colCount - colNo - 1);
				
			BOOL isTable = TRUE;
			xml_element* template = 0;
			int childCount = 0;

			xml_element* subnode = node->first_child;
			char* tagName = isSameFilter ? sameFilter8 : 0;	
			while (subnode) {
				childCount += subnode->key != 0 && (!isSameFilter || isSameFilter && strcmp(subnode->key, sameFilter8) == 0); 
				if (tagName && subnode->key && !isSameFilter) 
					isTable = isTable && strcmp(subnode->key, tagName) == 0;
					
				if (!template && subnode->key && strlen(subnode->key) > 0 && 
					(!isSameFilter  || 
					isSameFilter && stricmp(subnode->key, sameFilter8) == 0)) {
					tagName = subnode->key;
					template = subnode;
				}
										 
				subnode = subnode->next;
			}
			isTable = isTable && childCount > 1 || isSameFilter;			
			
			if (isTable) {				
				xml_attribute* attr = template->first_attribute;
				while (attr) {
					char* colName8 = calloc(strlen(attr->key) + 2, sizeof(char));
					sprintf(colName8, "@%s", attr->key);
					TCHAR* colName16 = utf8to16(colName8);
					ListView_AddColumn(hGridWnd, colName16);
					free(colName16);
					free(colName8);
														
					attr = attr->next;
				}	
				
				subnode = template->first_child;
				while (subnode) {
					if (subnode->key && strlen(subnode->key) > 0) {			
						TCHAR* colName16 = utf8to16(subnode->key);
						ListView_AddColumn(hGridWnd, colName16);
						free(colName16);					
					} 
					subnode = subnode->next;
				}
				if (isShowContent)
					ListView_AddColumn(hGridWnd, TEXT("#CONTENT"));
			} else {
				ListView_AddColumn(hGridWnd, TEXT("Key"));
				ListView_AddColumn(hGridWnd, TEXT("Value"));			
			}

			colCount = Header_GetItemCount(hHeader);
			int align = filterAlign == -1 ? ES_LEFT : filterAlign == 1 ? ES_RIGHT : ES_CENTER;
			for (int colNo = 0; colNo < colCount; colNo++) {
				// Use WS_BORDER to vertical text aligment
				HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, align | ES_AUTOHSCROLL | WS_CHILD | WS_TABSTOP | WS_BORDER,
					0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
				SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
				SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
			}
			
			
			// Cache data
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			ListView_SetItemCount(hGridWnd, 0);
			SetProp(hWnd, TEXT("CACHE"), 0);
		
			TCHAR*** cache = 0;
			xml_element** xmlnodes = 0;
			int rowCount = 0;
			int rowNo = 0;
			if (isTable) {
				rowCount = node->child_count;
				cache = calloc(rowCount, sizeof(TCHAR*));
				xmlnodes = calloc(rowCount, sizeof(xml_element*));
								
				xml_element* subnode = node->first_child;
				while (subnode) {
					if (!subnode->key || isSameFilter && strcmp(subnode->key, sameFilter8)) {
						subnode = subnode->next;
						continue;
					}
					
					cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
					xmlnodes[rowNo] = subnode;
					int colNo = 0;
					
					xml_attribute* attr = template->first_attribute;
					while (attr) {
						xml_attribute* sa = xml_find_attribute(subnode, attr->key);
						cache[rowNo][colNo] = utf8to16(sa ? sa->value : "N/A");
						
						colNo++;															
						attr = attr->next;
					}
					
					xml_element* tagNode = template->first_child;
					while (tagNode) {
						if (!tagNode->key) {
							tagNode = tagNode->next;
							continue;
						}
						
						xml_element* sn = xml_find_element(subnode, tagNode->key);
						char* content = sn ? xml_content(sn) : calloc(1, sizeof(char));
						cache[rowNo][colNo] = utf8to16(sn && !isEmpty(content) ? content : "N/A");
						free(content);
						
						colNo++;															
						tagNode = tagNode->next;
					}
					
					if (isShowContent) {
						char* content = subnode ? xml_content(subnode): calloc(1, sizeof(char));
						cache[rowNo][colNo] = utf8to16(subnode && !isEmpty(content) ? content : "");
						free(content);
					}
										
					rowNo++;
					subnode = subnode->next;
				}
			} else {
				rowCount = node->attribute_count + node->child_count + 1;
				cache = calloc(rowCount, sizeof(TCHAR*));
				xmlnodes = calloc(rowCount, sizeof(xml_element*));
				
				xml_attribute* attr = node->first_attribute;
				while (attr) {
					cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
					char* buf8 = calloc(strlen(attr->key) + 2, sizeof(char));
					sprintf(buf8, "@%s", attr->key);
					cache[rowNo][0] = utf8to16(buf8);
					free(buf8);
					cache[rowNo][1] = utf8to16(attr->value);

					rowNo++;				
					attr = attr->next;
				}
				
				BOOL isSpecial = node->key && (strncmp(node->key, "![CDATA[", 8) == 0 || strncmp(node->key, "!--", 3) == 0) || node->value;
				xml_element* subnode = isSpecial ? node : node->first_child;
				while (subnode) {
					if (!subnode->key && isEmpty(subnode->value)){
						subnode = subnode->next;
						continue;
					}
					
					cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
					xmlnodes[rowNo] = subnode;
					if (subnode->key) {
						if (strncmp(subnode->key, "![CDATA[", 8) == 0) {
							int len = strlen(subnode->key);
							char* buf = calloc(len + 1, sizeof(char));
							sprintf(buf, "%.*s", len - strlen("![CDATA[]]>"), subnode->key + 8);
							cache[rowNo][0] = utf8to16(XML_CDATA);
							cache[rowNo][1] = utf8to16(buf);
							free(buf);
						} else if (strncmp(subnode->key, "!--", 3) == 0) {
							int len = strlen(subnode->key);
							char* buf = calloc(len + 1, sizeof(char));
							sprintf(buf, "%.*s", len - strlen("!----"), subnode->key + 3);
							cache[rowNo][0] = utf8to16(XML_COMMENT);
							cache[rowNo][1] = utf8to16(buf);
							free(buf);
						} else {
							cache[rowNo][0] = utf8to16(subnode->key);
							
							char* content = xml_content(subnode);
							cache[rowNo][1] = utf8to16(!isEmpty(content) ? content : "");						
							free(content);
						}
					} else if (subnode->value && !isEmpty(subnode->value)) {
						cache[rowNo][0] = utf8to16(XML_TEXT);
						cache[rowNo][1] = utf8to16(subnode->value);					
					}
					
					rowNo++;
					subnode = isSpecial ? 0 : subnode->next;
				}
			}
							
			if (rowNo == 0) {
				if (cache)
					free(cache);
				if (xmlnodes)	
					free(xmlnodes);
			} else {
				cache = realloc(cache, rowNo * sizeof(TCHAR*));
				SetProp(hWnd, TEXT("CACHE"), cache);
				SetProp(hWnd, TEXT("XMLNODES"), xmlnodes);
			}
			
			SendMessage(hStatusWnd, SB_SETTEXT, SB_MODE, (LPARAM)(isSameFilter ? TEXT("  SAME") : isTable ? TEXT(" TABLE") : TEXT(" SINGLE")));
			TCHAR buf16[300];
			if (isSameFilter) {
				_sntprintf(buf16, 300, TEXT(" Shows only \"%ls\"-nodes"), sameFilter16);
			} else if (isTable) {
				_sntprintf(buf16, 300, TEXT(" Shows all child nodes"));
			} else {
				_sntprintf(buf16, 300, TEXT(" Shows the node"), sameFilter16);
			}
			SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)buf16);

			sameFilter16[0] = 0;
			free(sameFilter8);
			
			*pTotalRowCount = rowNo;
			SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
			SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
			SendMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_RESULTSET: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hGridWnd);

			ListView_SetItemCount(hGridWnd, 0);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
			int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
			BOOL isCaseSensitive = getStoredValue(TEXT("filter-case-sensitive"), 0);
			
			if (resultset)
				free(resultset);

			if (!cache)
				return 1;
				
			if (*pTotalRowCount == 0)	
				return 1;
				
			int colCount = Header_GetItemCount(hHeader);
			if (colCount == 0)
				return 1;
				
			BOOL* bResultset = (BOOL*)calloc(*pTotalRowCount, sizeof(BOOL));
			for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++)
				bResultset[rowNo] = TRUE;

			for (int colNo = 0; colNo < colCount; colNo++) {
				HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
				TCHAR filter[MAX_LENGTH];
				GetWindowText(hEdit, filter, MAX_LENGTH);
				int len = _tcslen(filter);
				if (len == 0)
					continue;

				for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
					if (!bResultset[rowNo])
						continue;

					TCHAR* value = cache[rowNo][colNo];
					if (len > 1 && (filter[0] == TEXT('<') || filter[0] == TEXT('>')) && isNumber(filter + 1)) {
						TCHAR* end = 0;
						double df = _tcstod(filter + 1, &end);
						double dv = _tcstod(value, &end);
						bResultset[rowNo] = (filter[0] == TEXT('<') && dv < df) || (filter[0] == TEXT('>') && dv > df);
					} else {
						bResultset[rowNo] = len == 1 ? hasString(value, filter, isCaseSensitive) :
							filter[0] == TEXT('=') && isCaseSensitive ? _tcscmp(value, filter + 1) == 0 :
							filter[0] == TEXT('=') && !isCaseSensitive ? _tcsicmp(value, filter + 1) == 0 :							
							filter[0] == TEXT('!') ? !hasString(value, filter + 1, isCaseSensitive) :
							filter[0] == TEXT('<') ? _tcscmp(value, filter + 1) < 0 :
							filter[0] == TEXT('>') ? _tcscmp(value, filter + 1) > 0 :
							hasString(value, filter, isCaseSensitive);
					}
				}
			}

			int rowCount = 0;
			resultset = (int*)calloc(*pTotalRowCount, sizeof(int));
			for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
				if (!bResultset[rowNo])
					continue;

				resultset[rowCount] = rowNo;
				rowCount++;
			}
			free(bResultset);

			if (rowCount > 0) {
				if (rowCount > *pTotalRowCount)
					MessageBeep(0);
				resultset = realloc(resultset, rowCount * sizeof(int));
				SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)resultset);
				int orderBy = *pOrderBy;

				if (orderBy) {
					int colNo = orderBy > 0 ? orderBy - 1 : - orderBy - 1;
					BOOL isBackward = orderBy < 0;
					
					BOOL isNum = TRUE;
					for (int i = 0; i < *pTotalRowCount && i <= 5; i++) 
						isNum = isNum && isNumber(cache[i][colNo]);
												
					if (isNum) {
						double* nums = calloc(*pTotalRowCount, sizeof(double));
						for (int i = 0; i < rowCount; i++)
							nums[resultset[i]] = _tcstod(cache[resultset[i]][colNo], NULL);

						mergeSort(resultset, (void*)nums, 0, rowCount - 1, isBackward, isNum);
						free(nums);
					} else {
						TCHAR** strings = calloc(*pTotalRowCount, sizeof(TCHAR*));
						for (int i = 0; i < rowCount; i++)
							strings[resultset[i]] = cache[resultset[i]][colNo];
						mergeSort(resultset, (void*)strings, 0, rowCount - 1, isBackward, isNum);
						free(strings);
					}
				}
			} else {
				SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)0);			
				free(resultset);
			}

			*pRowCount = rowCount;
			ListView_SetItemCount(hGridWnd, rowCount);
			InvalidateRect(hGridWnd, NULL, TRUE);

			TCHAR buf[255];
			_sntprintf(buf, 255, TEXT(" Rows: %i/%i"), rowCount, *pTotalRowCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)buf);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;
		
		case WMU_UPDATE_TEXT: {
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			char* data = (char*)GetProp(hWnd, TEXT("DATA"));
		
			HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);			
			xml_element* node = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
			if (!node)
				return 0;
			
			char* text8 = calloc(node->length + 1, sizeof(char));
			strncpy(text8, data + node->offset, node->length);
			
			if (*(int*)GetProp(hWnd, TEXT("ISFORMAT"))) {
				char* ftext8 = formatXML(text8);
				free(text8);
				text8 = ftext8;
			}
			
			TCHAR* text16 = utf8to16(text8);
			free(text8);
						
			LockWindowUpdate(hTextWnd);
			SendMessage(hTextWnd, EM_EXLIMITTEXT, 0, _tcslen(text16) + 1);
			SetWindowText(hTextWnd, text16);
			LockWindowUpdate(0);
			free(text16);		
			
			SendMessage(hWnd, WMU_UPDATE_HIGHLIGHT, 0, 0);	
		}
		break;
		
		case WMU_UPDATE_HIGHLIGHT: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);

			GETTEXTLENGTHEX gtl = {GTL_NUMBYTES, 0};
			int len = SendMessage(hTextWnd, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 1200);
			if (len < *(int*)GetProp(hWnd, TEXT("MAXHIGHLIGHTLENGTH"))) {
				TCHAR* text = calloc(len + sizeof(TCHAR), sizeof(char));
				GETTEXTEX gt = {0};
				gt.cb = len + sizeof(TCHAR);
				gt.flags = 0;
				gt.codepage = 1200;
				SendMessage(hTextWnd, EM_GETTEXTEX, (WPARAM)&gt, (LPARAM)text);
				LockWindowUpdate(hTextWnd);
				
				CHARFORMAT2 cf2 = {0};
				cf2.cbSize = sizeof(CHARFORMAT2) ;
				cf2.dwMask = CFM_COLOR;
				cf2.crTextColor = *(int*)GetProp(hWnd, TEXT("XMLTEXTCOLOR"));			
				SendMessage(hTextWnd, EM_SETSEL, 0, -1);
				SendMessage(hTextWnd, EM_SETCHARFORMAT, SCF_ALL, (LPARAM) &cf2);				
				
				highlightText(hTextWnd, text);
				free(text);
				SendMessage(hTextWnd, EM_SETSEL, 0, 0);
				LockWindowUpdate(0);
				InvalidateRect(hTextWnd, NULL, TRUE);
			}
		}
		break;

		case WMU_UPDATE_FILTER_SIZE: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hHeader, WM_SIZE, 0, 0);
			
			int* colOrder = calloc(colCount, sizeof(int));
			Header_GetOrderArray(hHeader, colCount, colOrder);

			for (int idx = 0; idx < colCount; idx++) {
				int colNo = colOrder[idx];
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left, h2, rc.right - rc.left, h2 + 1, SWP_NOZORDER);
			}

			free(colOrder);
		}
		break;
		
		case WMU_SET_HEADER_FILTERS: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int isFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW"));
			int colCount = Header_GetItemCount(hHeader);
			
			SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
			LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
			styles = isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR);
			SetWindowLongPtr(hHeader, GWL_STYLE, styles);
					
			for (int colNo = 0; colNo < colCount; colNo++) 		
				ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), isFilterRow ? SW_SHOW : SW_HIDE);

			// Bug fix: force Windows to redraw header
			SetWindowPos(hGridWnd, 0, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE);
			SendMessage(getMainWindow(hWnd), WM_SIZE, 0, 0);			

			if (isFilterRow)				
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);											
						
			SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		case WMU_AUTO_COLUMN_SIZE: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);

			for (int colNo = 0; colNo < colCount - 1; colNo++)
				ListView_SetColumnWidth(hGridWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

			if (colCount == 1 && ListView_GetColumnWidth(hGridWnd, 0) < 100)
				ListView_SetColumnWidth(hGridWnd, 0, 100);

			int maxWidth = getStoredValue(TEXT("max-column-width"), 300);
			if (colCount > 1) {
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
						ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
				}
			}

			// Fix last column
			if (colCount > 1) {
				int colNo = colCount - 1;
				ListView_SetColumnWidth(hGridWnd, colNo, LVSCW_AUTOSIZE);
				TCHAR name16[MAX_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, name16, MAX_LENGTH);
				
				SIZE s = {0};
				HDC hDC = GetDC(hHeader);
				HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)GetProp(hWnd, TEXT("FONT")));
				GetTextExtentPoint32(hDC, name16, _tcslen(name16), &s);
				SelectObject(hDC, hOldFont);
				ReleaseDC(hHeader, hDC);

				int w = s.cx + 12;
				if (ListView_GetColumnWidth(hGridWnd, colNo) < w)
					ListView_SetColumnWidth(hGridWnd, colNo, w);

				if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
					ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
			}

			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hGridWnd, NULL, TRUE);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;
		
		// wParam = rowNo, lParam = colNo
		case WMU_SET_CURRENT_CELL: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)0);
						
			int *pRowNo = (int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int *pColNo = (int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
			if (*pRowNo == wParam && *pColNo == lParam)
				return 0;
			
			RECT rc, rc2;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);			
			InvalidateRect(hGridWnd, &rc, TRUE);
			
			*pRowNo = wParam;
			*pColNo = lParam;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);
			InvalidateRect(hGridWnd, &rc, FALSE);
			
			GetClientRect(hGridWnd, &rc2);
			int w = rc.right - rc.left;
			int dx = rc2.right < rc.right ? rc.left - rc2.right + w : rc.left < 0 ? rc.left : 0;

			ListView_Scroll(hGridWnd, dx, 0);
			
			TCHAR buf[32] = {0};
			if (*pColNo != - 1 && *pRowNo != -1)
				_sntprintf(buf, 32, TEXT(" %i:%i"), *pRowNo + 1, *pColNo + 1);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, (LPARAM)buf);
		}
		break;			

		case WMU_RESET_CACHE: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));

			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			if (colCount > 0 && cache != 0) {
				for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
					if (cache[rowNo]) {
						for (int colNo = 0; colNo < colCount; colNo++)
							if (cache[rowNo][colNo])
								free(cache[rowNo][colNo]);

						free(cache[rowNo]);
					}
					cache[rowNo] = 0;
				}
				free(cache);
			}
			
			xml_element* xmlnodes = (xml_element*)GetProp(hWnd, TEXT("XMLNODES"));
			if (xmlnodes)
				free(xmlnodes);
			SetProp(hWnd, TEXT("XMLNODES"), 0);			

			int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
			if (resultset)
				free(resultset);
			SetProp(hWnd, TEXT("RESULTSET"), 0);
				
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			*pRowCount = 0;

			SetProp(hWnd, TEXT("CACHE"), 0);
			*pTotalRowCount = 0;
		}
		break;

		// wParam - size delta
		case WMU_SET_FONT: {
			int* pFontSize = (int*)GetProp(hWnd, TEXT("FONTSIZE"));
			if (*pFontSize + wParam < 10 || *pFontSize + wParam > 48)
				return 0;
			*pFontSize += wParam;
			DeleteFont(GetProp(hWnd, TEXT("FONT")));

			int weight = getStoredValue(TEXT("font-weight"), 0);
			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, weight < 0 || weight > 9 ? FW_DONTCARE : weight * 100, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			SendMessage(hTreeWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hTextWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hTabWnd, WM_SETFONT, (LPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);

			HWND hHeader = ListView_GetHeader(hGridWnd);
			for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
				SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);

			SetProp(hWnd, TEXT("FONT"), hFont);

			PostMessage(hWnd, WMU_UPDATE_HIGHLIGHT, 0, 0);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;
		
		case WMU_SET_THEME: {
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			BOOL isDark = *(int*)GetProp(hWnd, TEXT("DARKTHEME"));
			
			int textColor = !isDark ? getStoredValue(TEXT("text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("text-color-dark"), RGB(220, 220, 220));
			int backColor = !isDark ? getStoredValue(TEXT("back-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("back-color-dark"), RGB(32, 32, 32));
			int backColor2 = !isDark ? getStoredValue(TEXT("back-color2"), RGB(240, 240, 240)) : getStoredValue(TEXT("back-color2-dark"), RGB(52, 52, 52));
			int filterTextColor = !isDark ? getStoredValue(TEXT("filter-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("filter-text-color-dark"), RGB(255, 255, 255));
			int filterBackColor = !isDark ? getStoredValue(TEXT("filter-back-color"), RGB(240, 240, 240)) : getStoredValue(TEXT("filter-back-color-dark"), RGB(60, 60, 60));
			int currCellColor = !isDark ? getStoredValue(TEXT("current-cell-back-color"), RGB(70, 96, 166)) : getStoredValue(TEXT("current-cell-back-color-dark"), RGB(32, 62, 62));
			int selectionTextColor = !isDark ? getStoredValue(TEXT("selection-text-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("selection-text-color-dark"), RGB(220, 220, 220));
			int selectionBackColor = !isDark ? getStoredValue(TEXT("selection-back-color"), RGB(10, 36, 106)) : getStoredValue(TEXT("selection-back-color-dark"), RGB(72, 102, 102));			
			int splitterColor = !isDark ? getStoredValue(TEXT("splitter-color"), GetSysColor(COLOR_BTNFACE)) : getStoredValue(TEXT("splitter-color-dark"), GetSysColor(COLOR_BTNFACE));			
			
			*(int*)GetProp(hWnd, TEXT("TEXTCOLOR")) = textColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR")) = backColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR2")) = backColor2;			
			*(int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")) = filterTextColor;
			*(int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")) = filterBackColor;
			*(int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")) = currCellColor;
			*(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")) = selectionTextColor;
			*(int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")) = selectionBackColor;
			*(int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")) = splitterColor;
			
			*(int*)GetProp(hWnd, TEXT("XMLTEXTCOLOR")) = !isDark ? getStoredValue(TEXT("xml-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("xml-text-color-dark"), RGB(220, 220, 220));
			*(int*)GetProp(hWnd, TEXT("XMLTAGCOLOR")) = !isDark ? getStoredValue(TEXT("xml-tag-color"), RGB(0, 0, 128)) : getStoredValue(TEXT("xml-tag-color-dark"), RGB(0, 0, 255));			
			*(int*)GetProp(hWnd, TEXT("XMLSTRINGCOLOR")) = !isDark ? getStoredValue(TEXT("xml-string-color"), RGB(0, 128, 0)) : getStoredValue(TEXT("xml-string-color-dark"), RGB(0, 196, 0));
			*(int*)GetProp(hWnd, TEXT("XMLVALUECOLOR")) = !isDark ? getStoredValue(TEXT("xml-value-color"), RGB(0, 128, 128)) : getStoredValue(TEXT("xml-value-color-dark"), RGB(0, 196, 128));	
			*(int*)GetProp(hWnd, TEXT("XMLCDATACOLOR")) = !isDark ? getStoredValue(TEXT("xml-cdata-color"), RGB(220, 220, 220)) : getStoredValue(TEXT("xml-cdata-color-dark"), RGB(220, 220, 220));				
			*(int*)GetProp(hWnd, TEXT("XMLCOMMENTCOLOR")) = !isDark ? getStoredValue(TEXT("xml-comment-color"), RGB(220, 220, 128)) : getStoredValue(TEXT("xml-comment-color-dark"), RGB(220, 220, 128));							

			DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));			
			SetProp(hWnd, TEXT("BACKBRUSH"), CreateSolidBrush(backColor));
			SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(filterBackColor));
			SetProp(hWnd, TEXT("SPLITTERBRUSH"), CreateSolidBrush(splitterColor));			

			TreeView_SetTextColor(hTreeWnd, textColor);			
			TreeView_SetBkColor(hTreeWnd, backColor);

			ListView_SetTextColor(hGridWnd, textColor);			
			ListView_SetBkColor(hGridWnd, backColor);
			ListView_SetTextBkColor(hGridWnd, backColor);
						
			SendMessage(hTextWnd, EM_SETBKGNDCOLOR, 0, backColor);			
			
			SendMessage(hWnd, WMU_UPDATE_HIGHLIGHT, 0, 0);
			InvalidateRect(hWnd, NULL, TRUE);	
		}
		break;
		
		// wParam = tab index
		case WMU_SWITCH_TAB: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			TabCtrl_SetCurSel(hTabWnd, wParam);
			NMHDR Hdr = {hTabWnd, IDC_TAB, TCN_SELCHANGE};
			SendMessage(hWnd, WM_NOTIFY, IDC_TAB, (LPARAM)&Hdr);
		}
		break;
		
		case WMU_HOT_KEYS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			if (wParam == VK_TAB) {
				HWND hFocus = GetFocus();
				HWND wnds[1000] = {0};
				EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

				int no = 0;
				while(wnds[no] && wnds[no] != hFocus)
					no++;

				int cnt = no;
				while(wnds[cnt])
					cnt++;

				no += isCtrl ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isCtrl ? wnds[cnt - 1] : wnds[0]));
			}
			
			if (wParam == VK_F1) {
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/xmltab-wlx/wiki"), 0, 0 , SW_SHOW);
				return TRUE;
			}
			
			if (wParam == 0x20 && isCtrl) { // Ctrl + Space
				SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
				return TRUE;
			}
			
			if (wParam == VK_ESCAPE || wParam == VK_F11 ||
				wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || (isCtrl && wParam == 0x46) || // Ctrl + F
				((wParam >= 0x31 && wParam <= 0x38) && !getStoredValue(TEXT("disable-num-keys"), 0) && !isCtrl || // 1 - 8
				(wParam == 0x4E || wParam == 0x50) && !getStoredValue(TEXT("disable-np-keys"), 0) || // N, P
				wParam == 0x51 && getStoredValue(TEXT("exit-by-q"), 0)) && // Q
				GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT) {
				SetFocus(GetParent(hWnd));		
				keybd_event(wParam, wParam, KEYEVENTF_EXTENDEDKEY, 0);

				return TRUE;
			}

			if (isCtrl && wParam >= 0x30 && wParam <= 0x39 && !getStoredValue(TEXT("disable-num-keys"), 0)) {// 0-9
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));

				BOOL isCurrent = wParam == 0x30;
				int colNo = isCurrent ? *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) : wParam - 0x30 - 1;
				if (colNo < 0 || colNo > colCount - 1 || isCurrent && ListView_GetColumnWidth(hGridWnd, colNo) == 0)
					return FALSE;

				if (!isCurrent) {
					int* colOrder = calloc(colCount, sizeof(int));
					Header_GetOrderArray(hHeader, colCount, colOrder);

					int hiddenColCount = 0;
					for (int idx = 0; (idx < colCount) && (idx - hiddenColCount <= colNo); idx++)
						hiddenColCount += ListView_GetColumnWidth(hGridWnd, colOrder[idx]) == 0;

					colNo = colOrder[colNo + hiddenColCount];
					free(colOrder);
				}
				
				SendMessage(hWnd, WMU_SORT_COLUMN, colNo, 0);
				return TRUE;
			}			
			
			return FALSE;					
		}
		break;
		
		case WMU_HOT_CHARS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));

			unsigned char scancode = ((unsigned char*)&lParam)[2];
			UINT key = MapVirtualKey(scancode, MAPVK_VSC_TO_VK);

			return !_istprint(wParam) && (
				wParam == VK_ESCAPE || wParam == VK_F11 || wParam == VK_F1 ||
				wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7) ||
				wParam == VK_TAB || wParam == VK_RETURN ||
				isCtrl && key == 0x46 || // Ctrl + F
				getStoredValue(TEXT("exit-by-q"), 0) && key == 0x51 && GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT; // Q	
		}
		break;
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && SendMessage(getMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
		return 0;

	// Prevent beep
	if (msg == WM_CHAR && SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
		return 0;	

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_CTLCOLOREDIT) {
		HWND hMainWnd = getMainWindow(hWnd);
		SetBkColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
		SetTextColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR")));
		return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));	
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HWND hMainWnd = getMainWindow(hWnd);
			BOOL isDark = *(int*)GetProp(hMainWnd, TEXT("DARKTHEME")); 

			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			LineTo(hDC, rc.right - 1, rc.bottom - 1);
			
			if (isDark) {
				DeleteObject(hPen);
				hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
				SelectObject(hDC, hPen);
				
				MoveToEx(hDC, 0, 0, 0);
				LineTo(hDC, 0, rc.bottom);
				MoveToEx(hDC, 0, rc.bottom - 1, 0);
				LineTo(hDC, rc.right, rc.bottom - 1);
				MoveToEx(hDC, 0, rc.bottom - 2, 0);
				LineTo(hDC, rc.right, rc.bottom - 2);
			}
			
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;
		
		case WM_SETFOCUS: {
			SetProp(getMainWindow(hWnd), TEXT("LASTFOCUS"), hWnd);
		}
		break;
		
		case WM_KEYDOWN: {
			HWND hMainWnd = getMainWindow(hWnd);
			if (wParam == VK_RETURN) {
				SendMessage(hMainWnd, WMU_UPDATE_RESULTSET, 0, 0);
				return 0;			
			}
			
			if (SendMessage(hMainWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;
	
		// Prevent beep
		case WM_CHAR: {
			if (SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
				return 0;	
		}
		break;		

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewTab(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_NOTIFY) {
		NMHDR* pHdr = (LPNMHDR)lParam;	
		if (pHdr->idFrom == IDC_GRID && pHdr->code == (UINT)NM_CUSTOMDRAW)
			return SendMessage(GetParent(hWnd), msg, wParam, lParam);
	}
	
	if (msg == WM_KEYDOWN && SendMessage(getMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
		return 0;

	// Prevent beep
	if (msg == WM_CHAR && SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
		return 0;			

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewText(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == EM_SETZOOM)
		return 0;

	if (msg == WM_MOUSEWHEEL && LOWORD(wParam) == MK_CONTROL) {
		HWND hTabWnd = GetParent(hWnd);
		SendMessage(GetParent(hTabWnd), msg, wParam, lParam);
		return 0;
	}
	
	if (msg == WM_KEYDOWN && SendMessage(getMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
		return 0;

	// Prevent beep
	if (msg == WM_CHAR && SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
		return 0;			
		
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

HTREEITEM addNode(HWND hTreeWnd, HTREEITEM hParentItem, xml_element* node) {
	if (!node)
		return FALSE;
	
	HTREEITEM hItem = 0;	
	xml_element* subnode = node->first_child;
	if (node->key) {
		const char* value = node->value ? node->value : node->child_count == 1 && subnode && !subnode->key ? subnode->value : 0;
		char* buf8 = calloc(strlen(node->key) + (value ? strlen(value) : 0) + 10, sizeof(char));
		if (node->key && strncmp(node->key, "![CDATA[", 8) == 0) {
			sprintf(buf8, "%s (%.2fKb)", XML_CDATA, strlen(node->key)/1024.0);
		} else if (node->key && strncmp(node->key, "!--", 3) == 0) {
			sprintf(buf8, XML_COMMENT);
		} else if (!isEmpty(value)) {
			// trim value
			int l = 0, r = 0;
			int len = strlen(value);
			for (int i = 0; strchr(WHITESPACE, value[i]) && i < len; i++)
				l++;
			for (int i = len - 1; strchr(WHITESPACE, value[i]) && i > 0; i--)
				r++;				
			len -= l + r; 
			sprintf(buf8, "%s = %.*s", node->key, len, value + l);
		} else if (isEmpty(value) && node->child_count == 0 && getStoredValue(TEXT("show-empty"), 0)) {	
			sprintf(buf8, "%s = %s", node->key, "<none>");
		} else {
			sprintf(buf8, "%s", node->key);
		}
		
		TCHAR* name16 = utf8to16(buf8);
		hItem = TreeView_AddItem(hTreeWnd, name16, hParentItem, (LPARAM)node);	
		free(name16);
		free(buf8);
	} else if (node->value && !isEmpty(node->value)) {
		hItem = TreeView_AddItem(hTreeWnd, TEXT(XML_TEXT), hParentItem, (LPARAM)node);		
	}
	
	BOOL hasSubnode = FALSE;
	while (!hasSubnode && subnode) {
		hasSubnode = subnode->key != NULL;
		subnode = subnode->next;
	}
	if (hasSubnode)
		TreeView_AddItem(hTreeWnd, LOADING, hItem, (LPARAM)subnode);
		
	return hItem;	
}

void showNode(HWND hTreeWnd, xml_element* node) {
	if (node == NULL)
		return;
	
	xml_element* parent = node->parent;
	if (!parent)
		return;
		
	if (!parent->userdata)
		showNode(hTreeWnd, parent);
	
	TreeView_Expand(hTreeWnd, (HTREEITEM)parent->userdata, TVE_EXPAND);	
}

void highlightText (HWND hWnd, TCHAR* text) {
	BOOL inTag = FALSE;
	TCHAR q = 0; 	
	BOOL isValue = FALSE;

	HWND hMainWnd = getMainWindow(hWnd);
	COLORREF textColor = *(int*)GetProp(hMainWnd, TEXT("XMLTEXTCOLOR"));	
	COLORREF tagColor = *(int*)GetProp(hMainWnd, TEXT("XMLTAGCOLOR"));
	COLORREF stringColor = *(int*)GetProp(hMainWnd, TEXT("XMLSTRINGCOLOR"));	
	COLORREF valueColor = *(int*)GetProp(hMainWnd, TEXT("XMLVALUECOLOR"));	
	COLORREF cdataColor = *(int*)GetProp(hMainWnd, TEXT("XMLCDATACOLOR"));		
	COLORREF commentColor = *(int*)GetProp(hMainWnd, TEXT("XMLCOMMENTCOLOR"));
	BOOL useBold = getStoredValue(TEXT("font-use-bold"), 0);
	
	CHARFORMAT2 cf2 = {0};
	cf2.cbSize = sizeof(CHARFORMAT2) ;
	cf2.dwMask = CFM_COLOR | CFM_BOLD;	
		
	int len = _tcslen(text);
	int pos = 0;
	while (pos < len) {
		int start = pos;
		TCHAR c = text[pos];

		cf2.crTextColor = textColor;
		cf2.dwEffects = 0;

		if (c == TEXT('\'') || c == TEXT('"')) {
			TCHAR* p = _tcschr(text + pos + 1, c);
			if (p != NULL) {
				pos = p - text;
				cf2.crTextColor = stringColor;
			}
		} else if (c == TEXT('<') && _tcsncmp(text + pos, TEXT("<![CDATA["), 9) == 0) {
			TCHAR* p = _tcsstr(text + pos, TEXT("]]>")); 
			if (p != NULL) {
				pos = p - text + 2;
				cf2.crTextColor = cdataColor;
			}
		} else if (c == TEXT('<') && _tcsncmp(text + pos, TEXT("<!--"), 4) == 0) {
			TCHAR *p = _tcsstr(text + pos, TEXT("-->"));
			if (p != NULL) {
				pos = p - text + 2;
				cf2.crTextColor = commentColor;
			}
		} else if (c == TEXT('<')) {
			TCHAR* p = _tcspbrk(text + pos, TEXT(" \t\r\n>"));
			if (p != NULL) {
				pos = p - text; 
				cf2.crTextColor = tagColor;
				cf2.dwEffects = useBold ? CFM_BOLD : 0;
			}
		} else if (c == TEXT('>')) {
			if (pos > 0 && text[pos - 1] == TEXT('/'))
				start--;

			cf2.crTextColor = tagColor;
			cf2.dwEffects = useBold ? CFM_BOLD : 0;
		} else if (pos > 0 && text[pos - 1] == TEXT('>')) {
			TCHAR* p = _tcschr(text + pos + 1, TEXT('<'));
			if (p != NULL) {
				pos = p - text - 1; 	
				cf2.crTextColor = valueColor;
			}
		}

		SendMessage(hWnd, EM_SETSEL, start, pos + 1);
		SendMessage(hWnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM) &cf2);
		pos++;		
	}
}

char* formatXML(const char* data) {
	if (!data)
		return 0;

	int len = strlen(data);
	int bLen = len;
	char* buf = calloc(bLen, sizeof(char));

	BOOL inTag = FALSE;
	BOOL isAttr = FALSE;
	BOOL isValue = FALSE;
	char quote = 0;
	BOOL isOpenTag = FALSE;
	BOOL isCData = FALSE;	
	
	int bPos = 0;
	int pos = 0;
	int level = 0;
	while (pos < len) {
		char c = data[pos];
		char p = bPos > 0 ? buf[bPos - 1] : 0;
		char n = data[pos + 1];

		BOOL isSave = FALSE;
		BOOL isSpace = FALSE;
		BOOL isNun = strchr(WHITESPACE, c) != 0;
		BOOL isCDataStart = FALSE;

		if (1000 == 1000 && isCData) {
			isSave = TRUE; 	
			isCData = !(data[pos - 2] == ']' && data[pos - 1] == ']' && c == '>');
		} else if (quote && c != quote) {
			isSave = TRUE;
		} else if (quote && c == quote) {
			isSave = TRUE;
			quote = 0;
		} else if (c == '<') {
			inTag = TRUE;
			isSave = TRUE;
			isValue = FALSE;
			isOpenTag = n != '/';
			isCDataStart = strncmp(data + pos, "<![CDATA[", 9) == 0; 
			isCData = isCDataStart;
			if (!isOpenTag)
				level--;
		} else if (inTag && c == '>') {
			isSave = TRUE;
			inTag = FALSE;
			isAttr = FALSE;
			isValue = FALSE;
			if (p == '/' || p == '-')
				level--;
			quote = 0;
		} else if (inTag && c == '=') {
			isSave = TRUE;
			isSpace = TRUE;
		} else if (inTag && !isNun) {
			isSave = TRUE;
			isSpace = !isAttr && p != '<' && p != '/' || p == '=';
			isAttr = TRUE;
		} else if (inTag && isNun) {
			isAttr = FALSE;
		} else if (!inTag && p == '>') {
			isValue = FALSE;
			isSave = !isNun;
		} else if (!inTag && !isValue && !isNun) {
			isSave = TRUE;
			isValue = TRUE;
		} else if (!inTag && isValue && !isNun) {
			isSave = TRUE;
		} else if (!inTag && isValue && isNun) {
			int n = strspn(data + pos, WHITESPACE);
			BOOL isEmptyTail = data[pos + n] == '<';
			if (isEmptyTail) {
				isValue = FALSE;
			} else {
				isSave = TRUE;
			}
		}

		if (isSave && bPos > 0 && !quote && (c == '<' || isCDataStart || p == '>' && !isCData)) {
			buf[bPos] = '\n';
			bPos++;

			for (int l = 0; l < level; l++) {
				buf[bPos] = '\t';
				bPos++;
			}
		}

		if (isSpace) {
			buf[bPos] = ' ';
			bPos++;
		}

		if (isSave) {
			buf[bPos] = data[pos];
			bPos++;
			
			if (bLen - bPos < 1000) {
				bLen += 32000;
				buf = realloc(buf, bLen);
			}
		}

		if (c == '<' && isOpenTag && !isCData)
			level++;

		pos++;
	}
	buf[bPos] = 0;

	return buf;
}

HWND getMainWindow(HWND hWnd) {
	HWND hMainWnd = hWnd;
	while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
		hMainWnd = GetParent(hMainWnd);
	return hMainWnd;	
}

void setStoredValue(TCHAR* name, int value) {
	TCHAR buf[128];
	_sntprintf(buf, 128, TEXT("%i"), value);
	WritePrivateProfileString(APP_NAME, name, buf, iniPath);
}

int getStoredValue(TCHAR* name, int defValue) {
	TCHAR buf[128];
	return GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) ? _ttoi(buf) : defValue;
}

TCHAR* getStoredString(TCHAR* name, TCHAR* defValue) { 
	TCHAR* buf = calloc(256, sizeof(TCHAR));
	if (0 == GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) && defValue)
		_tcsncpy(buf, defValue, 255);
	return buf;	
}

int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam) {
	if (GetWindowLong(hWnd, GWL_STYLE) & WS_TABSTOP && IsWindowVisible(hWnd)) {
		int no = 0;
		HWND* wnds = (HWND*)lParam;
		while (wnds[no])
			no++;
		wnds[no] = hWnd;
	}

	return TRUE;
}

TCHAR* utf8to16(const char* in) {
	TCHAR *out;
	if (!in || strlen(in) == 0) {
		out = (TCHAR*)calloc (1, sizeof (TCHAR));
	} else  {
		DWORD size = MultiByteToWideChar(CP_UTF8, 0, in, -1, NULL, 0);
		out = (TCHAR*)calloc (size, sizeof (TCHAR));
		MultiByteToWideChar(CP_UTF8, 0, in, -1, out, size);
	}
	return out;
}

char* utf16to8(const TCHAR* in) {
	char* out;
	if (!in || _tcslen(in) == 0) {
		out = (char*)calloc (1, sizeof(char));
	} else  {
		int len = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, 0, 0);
		out = (char*)calloc (len, sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, in, -1, out, len, 0, 0);
	}
	return out;
}

int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords) {
	if (!text || !word)
		return -1;
		
	int res = -1;
	int tlen = _tcslen(text);
	int wlen = _tcslen(word);	
	if (!tlen || !wlen)
		return res;
	
	if (!isMatchCase) {
		TCHAR* ltext = calloc(tlen + 1, sizeof(TCHAR));
		_tcsncpy(ltext, text, tlen);
		text = _tcslwr(ltext);

		TCHAR* lword = calloc(wlen + 1, sizeof(TCHAR));
		_tcsncpy(lword, word, wlen);
		word = _tcslwr(lword);
	}

	if (isWholeWords) {
		for (int pos = 0; (res  == -1) && (pos <= tlen - wlen); pos++) 
			res = (pos == 0 || pos > 0 && !_istalnum(text[pos - 1])) && 
				!_istalnum(text[pos + wlen]) && 
				_tcsncmp(text + pos, word, wlen) == 0 ? pos : -1;
	} else {
		TCHAR* s = _tcsstr(text, word);
		res = s != NULL ? s - text : -1;
	}
	
	if (!isMatchCase) {
		free(text);
		free(word);
	}

	return res; 
}

int findString8(char* text, char* word, BOOL isMatchCase, BOOL isWholeWords) {
	if (!text || !word)
		return -1;
		
	int res = -1;
	int tlen = strlen(text);
	int wlen = strlen(word);	
	if (!tlen || !wlen)
		return res;
	
	if (!isMatchCase) {
		char* ltext = calloc(tlen + 1, sizeof(char));
		strncpy(ltext, text, tlen);
		text = strlwr(ltext);

		char* lword = calloc(wlen + 1, sizeof(char));
		strncpy(lword, word, wlen);
		word = strlwr(lword);
	}

	if (isWholeWords) {
		for (int pos = 0; (res  == -1) && (pos <= tlen - wlen); pos++) 
			res = (pos == 0 || pos > 0 && !isalnum(text[pos - 1])) && 
				!isalnum(text[pos + wlen]) && 
				strncmp(text + pos, word, wlen) == 0 ? pos : -1;
	} else {
		char* s = strstr(text, word);
		res = s != NULL ? s - text : -1;
	}
	
	if (!isMatchCase) {
		free(text);
		free(word);
	}

	return res; 
}

BOOL hasString (const TCHAR* str, const TCHAR* sub, BOOL isCaseSensitive) {
	BOOL res = FALSE;

	TCHAR* lstr = _tcsdup(str);
	_tcslwr(lstr);
	TCHAR* lsub = _tcsdup(sub);
	_tcslwr(lsub);
	res = isCaseSensitive ? _tcsstr(str, sub) != 0 : _tcsstr(lstr, lsub) != 0;
	free(lstr);
	free(lsub);

	return res;
};

TCHAR* extractUrl(TCHAR* data) {
	int len = data ? _tcslen(data) : 0;
	int start = len;
	int end = len;
	
	TCHAR* url = calloc(len + 10, sizeof(TCHAR));
	
	TCHAR* slashes = _tcsstr(data, TEXT("://"));
	if (slashes) {
		start = len - _tcslen(slashes);
		end = start + 3;
		for (; start > 0 && _istalpha(data[start - 1]); start--);
		for (; end < len && data[end] != TEXT(' ') && data[end] != TEXT('"') && data[end] != TEXT('\''); end++);
		_tcsncpy(url, data + start, end - start);
		
	} else if (_tcschr(data, TEXT('.'))) {
		_sntprintf(url, len + 10, TEXT("https://%ls"), data);
	}
	
	return url;
}

// https://stackoverflow.com/a/25023604/6121703
int detectCodePage(const unsigned char *data) {
	return strncmp(data, "\xEF\xBB\xBF", 3) == 0 ? CP_UTF8 : // BOM
		strncmp(data, "\xFE\xFF", 2) == 0 ? CP_UTF16BE : // BOM
		strncmp(data, "\xFF\xFE", 2) == 0 ? CP_UTF16LE : // BOM
		strncmp(data, "\x00\x3C", 2) == 0 ? CP_UTF16BE : // <		
		strncmp(data, "\x3C\x00", 2) == 0 ? CP_UTF16LE : // <
		isUtf8(data) ? CP_UTF8 :		
		CP_ACP;
}

void setClipboardText(const TCHAR* text) {
	int len = (_tcslen(text) + 1) * sizeof(TCHAR);
	HGLOBAL hMem =  GlobalAlloc(GMEM_MOVEABLE, len);
	memcpy(GlobalLock(hMem), text, len);
	GlobalUnlock(hMem);
	OpenClipboard(0);
	EmptyClipboard();
	SetClipboardData(CF_UNICODETEXT, hMem);
	CloseClipboard();
}

BOOL isNumber(const TCHAR* val) {
	int len = _tcslen(val);
	BOOL res = TRUE;
	int pCount = 0;
	for (int i = 0; res && i < len; i++) {
		pCount += val[i] == TEXT('.');
		res = _istdigit(val[i]) || val[i] == TEXT('.');
	}
	return res && pCount < 2;
}

BOOL isEmpty (const char* s) {
	BOOL res = TRUE;
	for (int i = 0; s && res && i < strlen(s); i++)
		res = strchr(WHITESPACE, s[i]) != NULL;	
		
	return res;	
}

// https://stackoverflow.com/a/1031773/6121703
BOOL isUtf8(const char * string) {
	if (!string)
		return FALSE;

	const unsigned char * bytes = (const unsigned char *)string;
	while (*bytes) {
		if((bytes[0] == 0x09 || bytes[0] == 0x0A || bytes[0] == 0x0D || (0x20 <= bytes[0] && bytes[0] <= 0x7E))) {
			bytes += 1;
			continue;
		}

		if (((0xC2 <= bytes[0] && bytes[0] <= 0xDF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF))) {
			bytes += 2;
			continue;
		}

		if ((bytes[0] == 0xE0 && (0xA0 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(((0xE1 <= bytes[0] && bytes[0] <= 0xEC) || bytes[0] == 0xEE || bytes[0] == 0xEF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(bytes[0] == 0xED && (0x80 <= bytes[1] && bytes[1] <= 0x9F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF))
		) {
			bytes += 3;
			continue;
		}

		if ((bytes[0] == 0xF0 && (0x90 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			((0xF1 <= bytes[0] && bytes[0] <= 0xF3) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			(bytes[0] == 0xF4 && (0x80 <= bytes[1] && bytes[1] <= 0x8F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF))
		) {
			bytes += 4;
			continue;
		}

		return FALSE;
	}

	return TRUE;
}

void mergeSortJoiner(int indexes[], void* data, int l, int m, int r, BOOL isBackward, BOOL isNums) {
	int n1 = m - l + 1;
	int n2 = r - m;
	
	int* L = calloc(n1, sizeof(int));
	int* R = calloc(n2, sizeof(int)); 
	
	for (int i = 0; i < n1; i++)
		L[i] = indexes[l + i];
	for (int j = 0; j < n2; j++)
		R[j] = indexes[m + 1 + j];
	
	int i = 0, j = 0, k = l;
	while (i < n1 && j < n2) {
		int cmp = isNums ? ((double*)data)[L[i]] <= ((double*)data)[R[j]] : _tcscmp(((TCHAR**)data)[L[i]], ((TCHAR**)data)[R[j]]) <= 0;
		if (isBackward)
			cmp = !cmp;
		
		if (cmp) {
			indexes[k] = L[i];
			i++;
		} else {
			indexes[k] = R[j];
			j++;
		}
		k++;
	}
	
	while (i < n1) {
		indexes[k] = L[i];
		i++;
		k++;
	}
	
	while (j < n2) {
		indexes[k] = R[j];
		j++;
		k++;
	}
	
	free(L);
	free(R);
}

void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums) {
	if (l < r) {
		int m = l + (r - l) / 2;
		mergeSort(indexes, data, l, m, isBackward, isNums);
		mergeSort(indexes, data, m + 1, r, isBackward, isNums);
		mergeSortJoiner(indexes, data, l, m, r, isBackward, isNums);
	}
}

HTREEITEM TreeView_AddItem (HWND hTreeWnd, TCHAR* caption, HTREEITEM parent, LPARAM lParam) {
	TVITEM tvi = {0};
	TVINSERTSTRUCT tvins = {0};
	tvi.mask = TVIF_TEXT | TVIF_PARAM;
	tvi.pszText = caption;
	tvi.cchTextMax = _tcslen(caption) + 1;
	tvi.lParam = lParam;

	tvins.item = tvi;
	tvins.hInsertAfter = TVI_FIRST;
	tvins.hParent = parent;
	return (HTREEITEM)SendMessage(hTreeWnd, TVM_INSERTITEM, 0, (LPARAM)(LPTVINSERTSTRUCT)&tvins);
};

int TreeView_GetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* buf, int maxLen) {
	TV_ITEM tv = {0};
	tv.mask = TVIF_TEXT;
	tv.hItem = hItem;
	tv.cchTextMax = maxLen;
	tv.pszText = buf;
	return TreeView_GetItem(hTreeWnd, &tv);
}

LPARAM TreeView_GetItemParam(HWND hTreeWnd, HTREEITEM hItem) {
	TV_ITEM tv = {0};
	tv.mask = TVIF_PARAM;
	tv.hItem = hItem;

	return TreeView_GetItem(hTreeWnd, &tv) ? tv.lParam : 0;
}

int TreeView_SetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* text) {
	TV_ITEM tv = {0};
	tv.mask = TVIF_TEXT;
	tv.hItem = hItem;
	tv.cchTextMax = _tcslen(text) + 1;
	tv.pszText = text;
	return TreeView_SetItem(hTreeWnd, &tv);
}

int ListView_AddColumn(HWND hListWnd, TCHAR* colName) {
	int colNo = Header_GetItemCount(ListView_GetHeader(hListWnd));
	LVCOLUMN lvc = {0};
	lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
	lvc.iSubItem = colNo;
	lvc.pszText = colName;
	lvc.cchTextMax = _tcslen(colName) + 1;
	lvc.cx = 100;
	return ListView_InsertColumn(hListWnd, colNo, &lvc);
}

int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax) {
	if (i < 0)
		return FALSE;

	TCHAR* buf = calloc(cchTextMax + 1, sizeof(TCHAR));

	HDITEM hdi = {0};
	hdi.mask = HDI_TEXT;
	hdi.pszText = buf;
	hdi.cchTextMax = cchTextMax;
	int rc = Header_GetItem(hWnd, i, &hdi);

	_tcsncpy(pszText, buf, cchTextMax);
	free(buf);
	return rc;
}

void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState) {
	MENUITEMINFO mii = {0};
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STATE;
	mii.fState = fState;
	SetMenuItemInfo(hMenu, wID, FALSE, &mii);
}