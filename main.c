#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
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
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 4
#define WMU_RESET_CACHE        WM_USER + 5
#define WMU_SET_FONT           WM_USER + 6
#define WMU_UPDATE_TEXT        WM_USER + 7
#define WMU_UPDATE_HIGHLIGHT   WM_USER + 8
#define WMU_SWITCH_TAB         WM_USER + 9

#define IDC_MAIN               100
#define IDC_TREE               101
#define IDC_TAB                102
#define IDC_GRID               103
#define IDC_TEXT               104
#define IDC_STATUSBAR          105
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROW           5001
#define IDM_COPY_TEXT          5002
#define IDM_SELECTALL          5003
#define IDM_FORMAT             5004
#define IDM_LOCATE             5005

#define SB_UNUSED              0
#define SB_CODEPAGE            1
#define SB_MODE                2
#define SB_ROW_COUNT           3
#define SB_CURRENT_ROW         4

#define SPLITTER_WIDTH         5
#define MAX_LENGTH             4096
#define MAX_HIGHLIGHT_LENGTH   64000
#define APP_NAME               TEXT("xmltab")
#define LOADING                TEXT("Loading...")

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
LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewText(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
HTREEITEM addNode(HWND hTreeWnd, HTREEITEM hParentItem, xml_element* val);
void highlightText(HWND hWnd, TCHAR* text);
char* formatXML(const char* data);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int detectCodePage(const unsigned char *data);
void setClipboardText(const TCHAR* text);
BOOL isNumber(const TCHAR* val);
BOOL isEmpty (const char* s);
BOOL isUtf8(const char * string);
HTREEITEM TreeView_AddItem (HWND hTreeWnd, TCHAR* caption, HTREEITEM parent, LPARAM lParam);
int TreeView_GetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* buf, int maxLen);
LPARAM TreeView_GetItemParam(HWND hTreeWnd, HTREEITEM hItem);
int TreeView_SetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* text);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	snprintf(DetectString, maxlen, "MULTIMEDIA & ext=\"XML\"");
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

int __stdcall ListSearchText(HWND hWnd, char* searchString, int searchParameter) {
	HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
	HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);

	DWORD len = MultiByteToWideChar(CP_ACP, 0, searchString, -1, NULL, 0);
	TCHAR* searchString16 = (TCHAR*)calloc (len, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, searchString, -1, searchString16, len);
	
	int spos = SendMessage(hTextWnd, EM_GETSEL, 0, 0);
	int mode = 0;	
	FINDTEXTEXW ft = {{HIWORD(spos), -1}, searchString16, {0, 0}};
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
		SendMessage(hTextWnd, EM_SETSEL, pos, pos + _tcslen(searchString16));
	else	
		MessageBeep(0);
	free(searchString16);	
	SetFocus(hTextWnd);		
	
	return 0;
}

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* filepath = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, filepath, size);

	struct _stat st = {0};
	if (_tstat(filepath, &st) != 0 || st.st_size > getStoredValue(TEXT("max-file-size"), 100000000) || st.st_size < 4)
		return 0;

	char* data = calloc(st.st_size + 1, sizeof(char));
	FILE *f = fopen(fileToLoad, "rb");
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
	if (!xml) {
		free(data);
		return 0;
	}

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	LoadLibrary(TEXT("msftedit.dll"));

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindowEx(WS_EX_CONTROLPARENT, WC_STATIC, TEXT("xml-wlx"), WS_CHILD | WS_VISIBLE | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("XML"), xml);
	SetProp(hMainWnd, TEXT("DATA"), data);	
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("RESULTSET"), 0);
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("COLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERPOSITION"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ISFORMAT"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));	
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("GRAYBRUSH"), CreateSolidBrush(GetSysColor(COLOR_BTNFACE)));
	SetProp(hMainWnd, TEXT("MAXHIGHLIGHTLENGTH"), calloc(1, sizeof(int)));	

	*(int*)GetProp(hMainWnd, TEXT("SPLITTERPOSITION")) = getStoredValue(TEXT("splitter-position"), 200);
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);
	*(int*)GetProp(hMainWnd, TEXT("ISFORMAT")) = getStoredValue(TEXT("format"), 1);
	*(int*)GetProp(hMainWnd, TEXT("MAXHIGHLIGHTLENGTH")) = getStoredValue(TEXT("max-highlight-length"), 64000);	

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE |  (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	int sizes[6] = {90, 150, 200, 400, 500, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 6, (LPARAM)&sizes);

	HWND hTreeWnd = CreateWindow(WC_TREEVIEW, NULL, WS_VISIBLE | WS_CHILD | TVS_HASBUTTONS | TVS_HASLINES | TVS_SHOWSELALWAYS | WS_TABSTOP,
		0, 0, 100, 100, hMainWnd, (HMENU)IDC_TREE, GetModuleHandle(0), NULL);
	SetProp(hTreeWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTreeWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	HWND hTabWnd = CreateWindow(WC_TABCONTROL, NULL, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | WS_TABSTOP, 100, 100, 100, 100,
		hMainWnd, (HMENU)IDC_TAB, GetModuleHandle(0), NULL);
	SetProp(hTabWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTabWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

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
	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, (tabNo == 0 ? WS_VISIBLE : 0) | WS_CHILD  | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SINGLESEL | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hTabWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));

	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);

	HWND hTextWnd = CreateWindowEx(0, TEXT("RICHEDIT50W"), NULL, (tabNo == 1 ? WS_VISIBLE : 0) | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN | WS_VSCROLL | WS_HSCROLL | WS_TABSTOP | ES_NOHIDESEL | ES_READONLY,
		0, 0, 100, 100, hTabWnd, (HMENU)IDC_TEXT, GetModuleHandle(0),  NULL);
	SetProp(hTextWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTextWnd, GWLP_WNDPROC, (LONG_PTR)cbNewText));
	TabCtrl_SetCurSel(hTabWnd, tabNo);

	HMENU hDataMenu = CreatePopupMenu();
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_ROW, TEXT("Copy row"));
	SetProp(hMainWnd, TEXT("DATAMENU"), hDataMenu);

	HMENU hTextMenu = CreatePopupMenu();
	AppendMenu(hTextMenu, MF_STRING, IDM_COPY_TEXT, TEXT("Copy"));
	AppendMenu(hTextMenu, MF_STRING, IDM_SELECTALL, TEXT("Select all"));
	AppendMenu(hTextMenu, MF_STRING, 0, NULL);
	AppendMenu(hTextMenu, MF_STRING | (*(int*)GetProp(hMainWnd, TEXT("ISFORMAT")) != 0 ? MF_CHECKED : 0), IDM_FORMAT, TEXT("Format"));
	AppendMenu(hTextMenu, MF_STRING, IDM_LOCATE, TEXT("Locate"));		
	SetProp(hMainWnd, TEXT("TEXTMENU"), hTextMenu);
	
	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WM_SIZE, 0, 0);
	
	SendMessage(hStatusWnd, SB_SETTEXT, SB_CODEPAGE, 
		(LPARAM)(cp == CP_UTF8 ? TEXT("    UTF-8") : cp == CP_UTF16LE ? TEXT(" UTF-16LE") : cp == CP_UTF16BE ? TEXT(" UTF-16BE") : TEXT("     ANSI")));	
	
	xml_element* node = xml->last_child;
	while (node) {
		HTREEITEM hItem = addNode(hTreeWnd, TVI_ROOT, node);	
		TreeView_Expand(hTreeWnd, hItem, TVE_EXPAND);		
		node = node->prev;
	}
	
	HTREEITEM hItem = TreeView_GetNextItem(hTreeWnd, TVI_ROOT, TVGN_CHILD);
	TreeView_Select(hTreeWnd, hItem, TVGN_CARET);
	SetFocus(hTreeWnd);

	return hMainWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("splitter-position"), *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("format"), *(int*)GetProp(hWnd, TEXT("ISFORMAT")));	
	setStoredValue(TEXT("tab-no"), TabCtrl_GetCurSel(GetDlgItem(hWnd, IDC_TAB)));

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	xml_free((xml_element*)GetProp(hWnd, TEXT("XML")));
	free((char*)GetProp(hWnd, TEXT("DATA")));
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("COLNO")));
	free((int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	free((int*)GetProp(hWnd, TEXT("ISFORMAT")));	
	free((int*)GetProp(hWnd, TEXT("MAXHIGHLIGHTLENGTH")));
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));		

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("GRAYBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("DATAMENU")));
	DestroyMenu(GetProp(hWnd, TEXT("TEXTMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("RESULTSET"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("COLNO"));
	RemoveProp(hWnd, TEXT("XML"));
	RemoveProp(hWnd, TEXT("DATA"));	
	RemoveProp(hWnd, TEXT("SPLITTERPOSITION"));
	RemoveProp(hWnd, TEXT("ISFORMAT"));
	RemoveProp(hWnd, TEXT("MAXHIGHLIGHTLENGTH"));	
	
	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));	
	RemoveProp(hWnd, TEXT("FONTSIZE"));	
	RemoveProp(hWnd, TEXT("GRAYBRUSH"));
	RemoveProp(hWnd, TEXT("DATAMENU"));
	RemoveProp(hWnd, TEXT("TEXTMENU"));

	DestroyWindow(hWnd);
	return;
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
			rc.right = rc.left + 5;
			FillRect(hDC, &rc, (HBRUSH)GetProp(hWnd, TEXT("GRAYBRUSH")));
			EndPaint(hWnd, &ps);

			return 0;
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;

		case WM_LBUTTONDOWN: {
			int x = GET_X_LPARAM(lParam);
			int pos = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			if (x >= pos || x <= pos + SPLITTER_WIDTH) {
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
			if (wParam != MK_LBUTTON || !GetProp(hWnd, TEXT("ISMOUSEDOWN")))
				return 0;

			DWORD x = GET_X_LPARAM(lParam);
			if (x > 0 && x < 32000)
				*(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")) = x;
			SendMessage(hWnd, WM_SIZE, 0, 0);
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
			if (wParam == VK_ESCAPE)
				SendMessage(GetParent(hWnd), WM_CLOSE, 0, 0);

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

				BOOL isBackward = HIWORD(GetKeyState(VK_CONTROL));
				no += isBackward ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isBackward ? wnds[cnt - 1] : wnds[0]));
			}
		}
		break;
		
		case WM_CONTEXTMENU: {
			POINT p = {LOWORD(lParam), HIWORD(lParam)};
			if (GetDlgCtrlID(WindowFromPoint(p)) == IDC_TEXT) {
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
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROW) {
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				if (rowNo != -1) {
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
					int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
					int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
					if (!resultset || rowNo >= rowCount)
						return 0;

					int colCount = Header_GetItemCount(hHeader);

					int startNo = cmd == IDM_COPY_CELL ? *(int*)GetProp(hWnd, TEXT("COLNO")) : 0;
					int endNo = cmd == IDM_COPY_CELL ? startNo + 1 : colCount;
					if (startNo > colCount || endNo > colCount)
						return 0;

					int len = 0;
					for (int colNo = startNo; colNo < endNo; colNo++) {
						int _rowNo = resultset[rowNo];
						len += _tcslen(cache[_rowNo][colNo]) + 1 /* column delimiter: TAB */;
					}

					TCHAR* buf = calloc(len + 1, sizeof(TCHAR));
					for (int colNo = startNo; colNo < endNo; colNo++) {
						int _rowNo = resultset[rowNo];
						_tcscat(buf, cache[_rowNo][colNo]);
						if (colNo != endNo - 1)
							_tcscat(buf, TEXT("\t"));
					}

					setClipboardText(buf);
					free(buf);
				}
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
							addNode(hTreeWnd, hItem, subnode);	
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
				NMLISTVIEW* pLV = (NMLISTVIEW*)lParam;
				int colNo = pLV->iSubItem + 1;
				int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
				int orderBy = *pOrderBy;
				*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
				SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				POINT p;
				GetCursorPos(&p);
				*(int*)GetProp(hWnd, TEXT("COLNO")) = ia->iSubItem;
				TrackPopupMenu(GetProp(hWnd, TEXT("DATAMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);

				TCHAR buf[255] = {0};
				int pos = ListView_GetNextItem(pHdr->hwndFrom, -1, LVNI_SELECTED);
				if (pos != -1)
					_sntprintf(buf, 255, TEXT(" %i"), pos + 1);
				SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_ROW, (LPARAM)buf);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43 && GetKeyState(VK_CONTROL)) // Ctrl + C
					SendMessage(hWnd, WM_COMMAND, IDM_COPY_ROW, 0);
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_DBLCLK) {
				NMITEMACTIVATE* ia = (NMITEMACTIVATE*) lParam;
				if (ia->iItem == -1)	
					return 0;
					
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				
				HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
				HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);
				hItem = TreeView_GetChild(hTreeWnd, hItem);
				for (int i = 0; i < resultset[ia->iItem]; i++)
					hItem = TreeView_GetNextSibling(hTreeWnd, hItem);
								
				if (hItem != 0)	{
					int rowNo = resultset[ia->iItem];
					TCHAR* key = cache[rowNo][0]; 
					if (_tcscmp(key, TEXT(XML_TEXT)) == 0 || _tcscmp(key, TEXT(XML_COMMENT)) == 0 || _tcscmp(key, TEXT(XML_CDATA)) == 0)
						SendMessage(hWnd, WMU_SWITCH_TAB, 1, 0);

					TreeView_SelectItem(hTreeWnd, hItem);
				} else {
					MessageBeep(0);	
				}
			}			

			if (pHdr->code == HDN_ITEMCHANGED && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(GetDlgItem(hWnd, IDC_TAB), IDC_GRID)))
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_GRID: {
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);

			HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);			
			xml_element* node = (xml_element*)TreeView_GetItemParam(hTreeWnd, hItem);
			if (!node)
				return 0;

			HWND hHeader = ListView_GetHeader(hGridWnd);
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
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
			char* tagName = 0;			
			while (isTable && subnode) {
				childCount += subnode->key != 0; 
				if (tagName && subnode->key) 
					isTable = strcmp(subnode->key, tagName) == 0;
					
				if (!tagName && subnode->key && strlen(subnode->key) > 0) {
					tagName = subnode->key;
					template = subnode;
				}
										 
				subnode = subnode->next;
			}
			isTable = isTable && childCount > 1;
			
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
				
				ListView_AddColumn(hGridWnd, TEXT("#CONTENT"));
			} else {
				ListView_AddColumn(hGridWnd, TEXT("Key"));
				ListView_AddColumn(hGridWnd, TEXT("Value"));			
			}

			colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++) {
				// Use WS_BORDER to vertical text aligment
				HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, ES_CENTER | ES_AUTOHSCROLL | WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_BORDER,
					0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
				SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
				SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
			}
			
			
			// Cache data
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			ListView_SetItemCount(hGridWnd, 0);
			SetProp(hWnd, TEXT("CACHE"), 0);
		
			TCHAR*** cache = 0;
			int rowCount = 0;
			int rowNo = 0;
			if (isTable) {
				rowCount = node->child_count;
				cache = calloc(rowCount, sizeof(TCHAR*));				
				
				xml_element* subnode = node->first_child;
				while (subnode) {
					if (!subnode->key) {
						subnode = subnode->next;
						continue;
					}
					
					cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
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
						cache[rowNo][colNo] = utf8to16(sn ? xml_content(sn) : "N/A");
						
						colNo++;															
						tagNode = tagNode->next;
					}
					cache[rowNo][colNo] = utf8to16(subnode ? xml_content(subnode) : "");										
					
					rowNo++;
					subnode = subnode->next;
				}
			} else {
				rowCount = node->attribute_count + node->child_count + 1;
				cache = calloc(rowCount, sizeof(TCHAR*));
				
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
				
				xml_element* subnode = node->first_child;
				while (subnode) {
					if (!subnode->key && isEmpty(subnode->value)){
						subnode = subnode->next;
						continue;
					}
					
					cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
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
							cache[rowNo][1] = utf8to16(xml_content(subnode));						
						}
					} else if (subnode->value && !isEmpty(subnode->value)) {
						cache[rowNo][0] = utf8to16(XML_TEXT);
						cache[rowNo][1] = utf8to16(subnode->value);					
					}
					
					rowNo++;
					subnode = subnode->next;
				}
			}	
							
			if (rowNo == 0) {
				if (cache)
					free(cache);
			} else {
				cache = realloc(cache, rowNo * sizeof(TCHAR*));
				SetProp(hWnd, TEXT("CACHE"), cache);
			}
			
			SendMessage(hStatusWnd, SB_SETTEXT, SB_MODE, (LPARAM)(isTable ? TEXT(" TABLE") : TEXT(" SINGLE")));								

			*pTotalRowCount = rowNo;
			SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
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
						bResultset[rowNo] = len == 1 ? _tcsstr(value, filter) != 0 :
							filter[0] == TEXT('=') ? _tcscmp(value, filter + 1) == 0 :
							filter[0] == TEXT('!') ? _tcsstr(value, filter + 1) == 0 :
							filter[0] == TEXT('<') ? _tcscmp(value, filter + 1) < 0 :
							filter[0] == TEXT('>') ? _tcscmp(value, filter + 1) > 0 :
							_tcsstr(value, filter) != 0;
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
					// Bubble-sort
					for (int i = 0; i < rowCount; i++) {
						for (int j = i + 1; j < rowCount; j++) {
							int a = resultset[i];
							int b = resultset[j];
								
							if(orderBy > 0 && _tcscmp(cache[a][orderBy - 1], cache[b][orderBy - 1]) > 0) {
								resultset[i] = b;
								resultset[j] = a;
							}

							if(orderBy < 0 && _tcscmp(cache[a][-orderBy - 1], cache[b][-orderBy - 1]) < 0) {
								resultset[i] = b;
								resultset[j] = a;
							}
						}
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
			for (int colNo = 0; colNo < colCount; colNo++) {
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left - (colNo > 0), h2, rc.right - rc.left + 1, h2 + 1, SWP_NOZORDER);
			}
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

			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
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
		
		// wParam = tab index
		case WMU_SWITCH_TAB: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			TabCtrl_SetCurSel(hTabWnd, wParam);
			NMHDR Hdr = {hTabWnd, IDC_TAB, TCN_SELCHANGE};
			SendMessage(hWnd, WM_NOTIFY, IDC_TAB, (LPARAM)&Hdr);
		}
		break;

	}
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		// Win10+ fix: draw an upper border
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetWindowRect(hWnd, &rc);
			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;

		case WM_CHAR: {
			return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);
		}
		break;

		case WM_KEYDOWN: {
			if (wParam == VK_RETURN) {
				HWND hHeader = GetParent(hWnd);
				HWND hGridWnd = GetParent(hHeader);
				HWND hTabWnd = GetParent(hGridWnd);
				HWND hMainWnd = GetParent(hTabWnd);
				SendMessage(hMainWnd, WMU_UPDATE_RESULTSET, 0, 0);
				
				return 0;			
			}
			
			if (wParam == VK_TAB || wParam == VK_ESCAPE)
				return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);
		}
		break;

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewText(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == EM_SETZOOM)
		return 0;

	if (msg == WM_MOUSEWHEEL && LOWORD(wParam) == MK_CONTROL) {
		HWND hTabWnd = GetParent(hWnd);
		SendMessage(GetParent(hTabWnd), msg, wParam, lParam);
		return 0;
	}
	
	if (msg == WM_KEYDOWN && (wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || HIWORD(GetKeyState(VK_CONTROL)) && wParam == 0x46)) { // Ctrl + F
		SendMessage(GetAncestor(hWnd, GA_ROOT), WM_KEYDOWN, wParam, lParam);
		return 0;
	}
	
	if (msg == WM_KEYDOWN && (wParam == VK_TAB || wParam == VK_ESCAPE)) { 
		return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);
	}	
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && (wParam == VK_TAB || wParam == VK_ESCAPE)) {
		HWND hMainWnd = hWnd;
		while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
			hMainWnd = GetParent(hMainWnd);
		SendMessage(hMainWnd, WM_KEYDOWN, wParam, lParam);
		return 0;
	}
	
	// Prevent beep
	if (msg == WM_CHAR && (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB))
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
			sprintf(buf8, "%s = %s", node->key, value);
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

void highlightText (HWND hWnd, TCHAR* text) {
	BOOL inTag = FALSE;
	TCHAR q = 0; 	
	BOOL isValue = FALSE;
	
	CHARFORMAT2 cf2 = {0};
	cf2.cbSize = sizeof(CHARFORMAT2) ;
	cf2.dwMask = CFM_COLOR | CFM_BOLD;	
		
	int len = _tcslen(text);
	int pos = 0;
	while (pos < len) {
		int start = pos;
		TCHAR c = text[pos];

		cf2.crTextColor = RGB(0, 0, 0);
		cf2.dwEffects = 0;

		if (c == TEXT('\'') || c == TEXT('"')) {
			TCHAR* p = _tcschr(text + pos + 1, c);
			if (p != NULL) {
				pos = p - text;
				cf2.crTextColor = RGB(0, 200, 0);
			}
		} else if (c == TEXT('<') && _tcsncmp(text + pos, TEXT("<![CDATA["), 9) == 0) {
			TCHAR* p = _tcsstr(text + pos, TEXT("]]>")); 
			if (p != NULL) {
				pos = p - text + 2;
				cf2.crTextColor = RGB(196, 196, 196);
			}
		} else if (c == TEXT('<') && _tcsncmp(text + pos, TEXT("<!--"), 4) == 0) {
			TCHAR *p = _tcsstr(text + pos, TEXT("-->"));
			if (p != NULL) {
				pos = p - text + 2;
				cf2.crTextColor = RGB(126, 0, 0);
			}
		} else if (c == TEXT('<')) {
			TCHAR* p = _tcspbrk(text + pos, TEXT(" \t\r\n>"));
			if (p != NULL) {
				pos = p - text; 
				cf2.crTextColor = RGB(0, 0, 128);
				cf2.dwEffects = CFM_BOLD;
			}
		} else if (c == TEXT('>')) {
			if (pos > 0 && text[pos - 1] == TEXT('/'))
				start--;

			cf2.crTextColor = RGB(0, 0, 128);
			cf2.dwEffects = CFM_BOLD;
		} else if (pos > 0 && text[pos - 1] == TEXT('>')) {
			TCHAR* p = _tcschr(text + pos + 1, TEXT('<'));
			if (p != NULL) {
				pos = p - text - 1; 	
				cf2.crTextColor = RGB(255, 128, 0);
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
		BOOL isNun = strchr(" \t\r\n", c) != 0;
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
			int n = strspn(data + pos, " \t\r\n");
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
			
			if (bLen - bPos < 100) {
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
		res = strchr(" \t\n\r", s[i]) != NULL;	
		
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