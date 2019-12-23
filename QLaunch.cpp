/*
   Quero Launcher
   http://www.quero.at/
   Copyright 2006 Viktor Krammer

   This file is part of Quero Launcher.

   Quero Launcher is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   Quero Launcher is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with Quero Toolbar.  If not, see <http://www.gnu.org/licenses/>.
*/

// Quero Launcher
// QLaunch.cpp

#include "stdafx.h"

#ifndef COMPILE_FOR_WIN9X
	#define _UNICODE
#endif
#include <tchar.h>
#include <wchar.h>
#include <atlbase.h>
#include <mshtml.h>
#include <exdisp.h>
#include <shlwapi.h>
#include <strsafe.h>
#include <shlguid.h>
#include <urlhist.h>
#include <shlobj.h>
#include <wininet.h>
//#include <msiehost.h>

// From msiehost.h
EXTERN_C const GUID CGID_InternetExplorer;
#define IECMDID_CLEAR_AUTOCOMPLETE_FOR_FORMS        0
#define IECMDID_SETID_AUTOCOMPLETE_FOR_FORMS        1
#define IECMDID_BEFORENAVIGATE_GETSHELLBROWSE     2
#define IECMDID_BEFORENAVIGATE_DOEXTERNALBROWSE   3
#define IECMDID_BEFORENAVIGATE_GETIDLIST          4
// Values for first parameter of IEID_CLEAR_AUTOCOMPLETE_FOR_FORMS
#define IECMDID_ARG_CLEAR_FORMS_ALL                 0
#define IECMDID_ARG_CLEAR_FORMS_ALL_BUT_PASSWORDS   1
#define IECMDID_ARG_CLEAR_FORMS_PASSWORDS_ONLY      2


#define EXIT_SUCCESS 0
#define EXIT_ERROR 1

#define MAX_URL_LEN 2048
#define MAX_INSTANCES 16
#define MAX_OPTION_LEN 32

#define MIN_WIDTH 0x80
#define MAX_WIDTH 0x2000
#define MIN_HEIGHT 0x20
#define MAX_HEIGHT 0x2000

#define OPTION_DEFAULT 0x0
#define OPTION_STARTPROCESS 0x1
#define OPTION_ALWAYSONTOP 0x2
#define OPTION_NOTITLEBAR 0x4
#define OPTION_KIOSK 0x8
#define OPTION_THEATER 0x10
#define OPTION_LOCATION_SET 0x20
#define OPTION_LOCATION 0x40
#define OPTION_MENUBAR_SET 0x80
#define OPTION_MENUBAR 0x100
#define OPTION_STATUS_SET 0x200
#define OPTION_STATUS 0x400
#define OPTION_TOOLBAR_SET 0x800
#define OPTION_TOOLBAR 0x1000
#define OPTION_SILENT_SET 0x2000
#define OPTION_SILENT 0x4000
#define OPTION_RESIZEABLE_SET 0x8000
#define OPTION_RESIZEABLE 0x10000
#define OPTION_QTOOLBAR_SET 0x20000
#define OPTION_QTOOLBAR 0x40000
#define OPTION_EXPLORERBAR_SET 0x80000

#define EXPLORERBAR_NONE 0
#define EXPLORERBAR_SEARCH 1
#define EXPLORERBAR_FAVORITES 2
#define EXPLORERBAR_HISTORY 3
#define EXPLORERBAR_CHANNELS 4
#define EXPLORERBAR_FOLDER 5
#define MAX_EXPLORERBARS 6

#define WAITFORSTARTUP 300
#define MAX_RETRIES (30 /*Seconds*/ * (1000/WAITFORSTARTUP))

#define CLEAR_ALL ~0
#define CLEAR_HISTORY 0x1
#define CLEAR_COOKIES 0x2
#define CLEAR_CACHE 0x4
#define CLEAR_AUTOCOMPLETE_FORMS 0x8
#define CLEAR_AUTOCOMPLETE_PWDS 0x10
#define CLEAR_FAVORITES 0x20

#define MAX_INET_CACHE_ENTRY_SIZE 2048

#define MAX_UA_LEN 2048
#define UA_PREFIX 0
#define UA_COMPATIBLE 1
#define UA_VERSION 2
#define UA_PLATFORM 3
#define UA_NFIELDS 4

// Instance Properties

typedef struct InstanceProperties
{
	int Left,Top,Width,Height;
	bool bLeft,bTop,bWidth,bHeight,bOpen;
	int ShowCommand;
	UINT Options;
	UINT ExplorerBar;
	HWND hWnd;
	PROCESS_INFORMATION ProcessInfo;
	TCHAR Open[3*MAX_URL_LEN];
} InstancePropertie;

// Explorer bars

typedef struct ExplorerBar
{
	TCHAR* Name;
	TCHAR* CLSID;
} ExplorerBar;

// Prototypes

BOOL CALLBACK FoundThreadHwnd(HWND h,LPARAM l);
void PrepareWindow(HWND hwnd,InstanceProperties *pInstance);
void ParseURLFile(TCHAR *pOption);
void StripQutoes(TCHAR *pDst,TCHAR *pSrc);
void DeleteTraces(UINT clear);
bool DeleteUrlCache(UINT type);
bool EmptyDirectory(LPCTSTR szPath,bool bDeleteDesktopIni);
void WriteUserAgent(TCHAR CustomUserAgent[UA_NFIELDS][MAX_UA_LEN],bool bCustomUserAgent[UA_NFIELDS]);

// Global variables

InstanceProperties Instances[MAX_INSTANCES];
ExplorerBar ExplorerBars[MAX_EXPLORERBARS]={
	{_T("none"),NULL},
	{_T("Search"),_T("{30D02401-6A81-11D0-8274-00C04FD5AE38}")},
	{_T("Favorites"),_T("{EFA24E61-B078-11D0-89E4-00C04FC9E26E}")},
	{_T("History"),_T("{EFA24E62-B078-11D0-89E4-00C04FC9E26E}")},				   
	{_T("Channels"),_T("{EFA24E63-B078-11D0-89E4-00C04FC9E26E}")},
	{_T("Folders"),_T("{EFA24E64-B078-11d0-89E4-00C04FC9E26E}")}
};

// IE7 constants
#define navOpenInNewTab 0x0800
#define navOpenBackgroundTab 0x1000

#define MAX_IE_TABS 50

int Focus;
bool bFocus;

int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
	IWebBrowser2 *pIE = NULL;
	HRESULT hr;
	HWND hwnd,parenthwnd;
	int ExitCode=EXIT_SUCCESS;

	int n,i;
	int nStartProcess;

	TCHAR Option[MAX_URL_LEN];
	TCHAR Value[MAX_URL_LEN];
	int iValue;
	bool IsInstance;
	bool IsGlobalTaskPresent; // deleteall etc.
	bool IsRelativeValue;
	bool IsQuoted;
	bool IsFirstProperty;
	UINT DeleteTracesOptions;
	TCHAR CustomUserAgent[UA_NFIELDS][MAX_UA_LEN];
	bool bCustomUserAgent[UA_NFIELDS];
	bool IsCustomUserAgentSpecified;

	InstanceProperties *pCurrentInstance;

	TCHAR ch;

	TCHAR *pCmdLine;

	CoInitialize(NULL);

#ifdef COMPILE_FOR_WIN9X
	pCmdLine=lpCmdLine;
#else
	pCmdLine=GetCommandLine();

	// Skip command name

	bool quotes=false;

	ch=*pCmdLine;

	while(ch && (quotes || !_istspace(ch)))
	{
		if(ch==L'"') quotes=!quotes;
		pCmdLine++;
		ch=*pCmdLine;
	}
#endif

//	pCmdLine=_T("top=0,left=0,open=about:blank top=0,left=890";
//	pCmdLine=_T("D:\\Data\\quero\\QLaunch\\Release\\QLaunch.exe start=e:\\soft\\firefox\\firefox.exe,left=0,top=0 x start=E:\\Soft\\Opera\\Opera.exe,top=30%,left=0";
//	pCmdLine=_T("\"C:\\Documents and Settings\\Viktor\\Desktop\\test.html\"";
//	pCmdLine=_T("deleteforms";


	// Get the width and height of the primary display

	RECT PrimaryDisplay;
	int DisplayWidth,DisplayHeight;
				
	SystemParametersInfo(SPI_GETWORKAREA,0,&PrimaryDisplay,0);
	DisplayWidth=PrimaryDisplay.right-PrimaryDisplay.left;
	DisplayHeight=PrimaryDisplay.bottom-PrimaryDisplay.top;

	n=0;
	nStartProcess=0;
	Focus=0;
	bFocus=false;
	IsGlobalTaskPresent=false;
	DeleteTracesOptions=0;

	pCurrentInstance=Instances;

	while(_istspace(*pCmdLine)) pCmdLine++; // Skip leading whitespaces

	while(n<MAX_INSTANCES && *pCmdLine)
	{
		pCurrentInstance->bLeft=pCurrentInstance->bTop=pCurrentInstance->bWidth=pCurrentInstance->bHeight=pCurrentInstance->bOpen=false;
		pCurrentInstance->Options=OPTION_DEFAULT;
		pCurrentInstance->ShowCommand=SW_SHOWNORMAL;
		pCurrentInstance->ProcessInfo.hProcess=NULL;

		// Parse window paramters

		IsInstance=false;
		IsFirstProperty=true;
		ch=*pCmdLine;
		do
		{
			i=0;

			IsQuoted=false;
			while(ch && (IsQuoted || (ch!=L'=' && ch!=',' && ch!=' ')))
			{
				if(ch==L'"') IsQuoted=!IsQuoted;
				if(i<MAX_URL_LEN-1) Option[i++]=ch;
				pCmdLine++;
				ch=*pCmdLine;
			}
			Option[i]=0;

			if(ch=='=')
			{
				i=0;
				pCmdLine++;
				ch=*pCmdLine;
				if(ch==L'"') // Skip leading quote
				{
					pCmdLine++;
					ch=*pCmdLine;
					IsQuoted=true;
				}
				else IsQuoted=false;
				
				while(ch && (IsQuoted?ch!=L'"':(ch!=L' ' && ch!=',')))
				{
					if(ch==L'\\' && IsQuoted && ((*(pCmdLine+1))==L'"' || (*(pCmdLine+1))==L'\\'))
					{
						pCmdLine++;
						ch=*pCmdLine;
					}
					if(i<MAX_URL_LEN-1) Value[i++]=ch;
					pCmdLine++;
					ch=*pCmdLine;
				}
				if(IsQuoted && ch) // Skip ending quote
				{
					pCmdLine++;
					ch=*pCmdLine;
				}
				if(Value[i-1]==L'%')
				{
					i--;
					IsRelativeValue=true;
				}
				else IsRelativeValue=false;
				Value[i]=0;
			}
			else Value[0]=0;

			if(!_tcsicmp(Option,_T("left")))
			{
				if(StrToIntEx(Value,STIF_DEFAULT,&iValue)==TRUE)
				{
					if(IsRelativeValue) iValue=PrimaryDisplay.left+(int)(DisplayWidth*(iValue*0.01));
					if(iValue>=-MAX_WIDTH && iValue<=MAX_WIDTH)
					{
						pCurrentInstance->Left=iValue;
						pCurrentInstance->bLeft=true;
					}
					IsInstance=true;
				}
			}
			else if(!_tcsicmp(Option,_T("top")))
			{
				if(StrToIntEx(Value,STIF_DEFAULT,&iValue)==TRUE)
				{
					if(IsRelativeValue) iValue=PrimaryDisplay.top+(int)(DisplayHeight*(iValue*0.01));
					if(iValue>=-MAX_HEIGHT && iValue<=MAX_HEIGHT)
					{
						pCurrentInstance->Top=iValue;
						pCurrentInstance->bTop=true;
					}
					IsInstance=true;
				}
			}
			else if(!_tcsicmp(Option,_T("width")))
			{
				if(StrToIntEx(Value,STIF_DEFAULT,&iValue)==TRUE)
				{
					if(IsRelativeValue) iValue=(int)(DisplayWidth*(iValue*0.01));
					if(iValue>=MIN_WIDTH && iValue<=MAX_WIDTH)
					{
						pCurrentInstance->Width=iValue;
						pCurrentInstance->bWidth=true;
					}
					IsInstance=true;
				}
			}
			else if(!_tcsicmp(Option,_T("height")))
			{
				if(StrToIntEx(Value,STIF_DEFAULT,&iValue)==TRUE)
				{
					if(IsRelativeValue) iValue=(int)(DisplayHeight*(iValue*0.01));
					if(iValue>=MIN_HEIGHT && iValue<=MAX_HEIGHT)
					{
						pCurrentInstance->Height=iValue;
						pCurrentInstance->bHeight=true;
					}
					IsInstance=true;
				}
			}
			else if(!_tcsicmp(Option,_T("open")))
			{
				if(Value[0])
				{
					StringCbCopy(pCurrentInstance->Open,sizeof pCurrentInstance->Open,Value);
					pCurrentInstance->bOpen=true;
				}
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("start")))
			{
				StringCbCopy(pCurrentInstance->Open,sizeof pCurrentInstance->Open,Value);
				pCurrentInstance->Options|=OPTION_STARTPROCESS;
				nStartProcess++;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("focus")))
			{
				Focus=n;
				bFocus=true;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("maximized")))
			{
				pCurrentInstance->ShowCommand=SW_MAXIMIZE;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("minimized")))
			{
				pCurrentInstance->ShowCommand=SW_SHOWMINNOACTIVE;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("alwaysontop")))
			{
				pCurrentInstance->Options|=OPTION_ALWAYSONTOP;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("notitlebar")))
			{
				pCurrentInstance->Options|=OPTION_NOTITLEBAR;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("location")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=OPTION_LOCATION|OPTION_LOCATION_SET;
				else if(!_tcsicmp(Value,_T("no"))) iValue=OPTION_LOCATION_SET;
				else iValue=0;

				pCurrentInstance->Options|=iValue;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("menubar")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=OPTION_MENUBAR|OPTION_MENUBAR_SET;
				else if(!_tcsicmp(Value,_T("no"))) iValue=OPTION_MENUBAR_SET;
				else iValue=0;

				pCurrentInstance->Options|=iValue;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("resizeable")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=OPTION_RESIZEABLE|OPTION_RESIZEABLE_SET;
				else if(!_tcsicmp(Value,_T("no"))) iValue=OPTION_RESIZEABLE_SET;
				else iValue=0;

				pCurrentInstance->Options|=iValue;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("silent")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=OPTION_SILENT|OPTION_SILENT_SET;
				else if(!_tcsicmp(Value,_T("no"))) iValue=OPTION_SILENT_SET;
				else iValue=0;

				pCurrentInstance->Options|=iValue;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("status")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=OPTION_STATUS|OPTION_STATUS_SET;
				else if(!_tcsicmp(Value,_T("no"))) iValue=OPTION_STATUS_SET;
				else iValue=0;

				pCurrentInstance->Options|=iValue;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("toolbar")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=OPTION_TOOLBAR|OPTION_TOOLBAR_SET;
				else if(!_tcsicmp(Value,_T("no"))) iValue=OPTION_TOOLBAR_SET;
				else iValue=0;

				pCurrentInstance->Options|=iValue;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("qtoolbar")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=OPTION_QTOOLBAR|OPTION_QTOOLBAR_SET;
				else if(!_tcsicmp(Value,_T("no"))) iValue=OPTION_QTOOLBAR_SET;
				else iValue=0;

				pCurrentInstance->Options|=iValue;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("explorerbar")))
			{
				i=0;
				while(i<MAX_EXPLORERBARS)
				{
					if(!_tcsicmp(Value,ExplorerBars[i].Name))
					{
						pCurrentInstance->Options|=OPTION_EXPLORERBAR_SET;
						pCurrentInstance->ExplorerBar=i;

						IsInstance=true;
						break;
					}
					i++;
				}
			}
			else if(!_tcsicmp(Option,_T("kiosk")))
			{
				pCurrentInstance->Options|=OPTION_KIOSK;
				pCurrentInstance->ShowCommand=SW_SHOW;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("theater")))
			{
				pCurrentInstance->Options|=OPTION_THEATER;
				pCurrentInstance->ShowCommand=SW_SHOW;
				IsInstance=true;
			}
			else if(!_tcsicmp(Option,_T("deleteall")))
			{
				DeleteTracesOptions|=CLEAR_ALL;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("deletehistory")))
			{
				DeleteTracesOptions|=CLEAR_HISTORY;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("deletecookies")))
			{
				DeleteTracesOptions|=CLEAR_COOKIES;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("deletecache")))
			{
				DeleteTracesOptions|=CLEAR_CACHE;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("deleteforms")))
			{
				DeleteTracesOptions|=CLEAR_AUTOCOMPLETE_FORMS;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("deletepasswords")))
			{
				DeleteTracesOptions|=CLEAR_AUTOCOMPLETE_PWDS;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("deletefavorites")))
			{
				DeleteTracesOptions|=CLEAR_FAVORITES;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("itbar7position")))
			{
				if(StrToIntEx(Value,STIF_DEFAULT,&iValue)==TRUE)
				{
					SHSetValue(HKEY_CURRENT_USER,_T("Software\\Microsoft\\Internet Explorer\\Toolbar\\WebBrowser"),_T("ITBar7Position"),REG_DWORD,&iValue,4);
					IsGlobalTaskPresent=true;
				}
			}
			else if(!_tcsicmp(Option,_T("commandbar")))
			{
				if(!_tcsicmp(Value,_T("yes"))) SHDeleteValue(HKEY_CURRENT_USER,_T("Software\\Microsoft\\Internet Explorer\\CommandBar"),_T("Enabled"));
				else if(!_tcsicmp(Value,_T("no")))
				{
					iValue=0;
					SHSetValue(HKEY_CURRENT_USER,_T("Software\\Microsoft\\Internet Explorer\\CommandBar"),_T("Enabled"),REG_DWORD,&iValue,4);
				}

				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("searchbox")))
			{
				if(!_tcsicmp(Value,_T("yes"))) SHDeleteValue(HKEY_CURRENT_USER,_T("Software\\Policies\\Microsoft\\Internet Explorer\\Infodelivery\\Restrictions"),_T("NoSearchBox"));
				else if(!_tcsicmp(Value,_T("no")))
				{
					iValue=1;
					SHSetValue(HKEY_CURRENT_USER,_T("Software\\Policies\\Microsoft\\Internet Explorer\\Infodelivery\\Restrictions"),_T("NoSearchBox"),REG_DWORD,&iValue,4);
				}

				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("tabbedbrowsing")))
			{
				if(!_tcsicmp(Value,_T("yes"))) iValue=1;
				else if(!_tcsicmp(Value,_T("no"))) iValue=0;
				else iValue=-1;

				if(iValue!=-1)
				{
					SHSetValue(HKEY_CURRENT_USER,_T("Software\\Microsoft\\Internet Explorer\\TabbedBrowsing"),_T("Enabled"),REG_DWORD,&iValue,4);
					IsGlobalTaskPresent=true;
				}
			}
			else if(!_tcsicmp(Option,_T("ua-prefix")))
			{
				StringCchCopy(CustomUserAgent[UA_PREFIX],MAX_UA_LEN,Value);
				bCustomUserAgent[UA_PREFIX]=true;
				IsCustomUserAgentSpecified=true;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("ua-compatible")))
			{
				StringCchCopy(CustomUserAgent[UA_COMPATIBLE],MAX_UA_LEN,Value);
				bCustomUserAgent[UA_COMPATIBLE]=true;
				IsCustomUserAgentSpecified=true;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("ua-version")))
			{
				StringCchCopy(CustomUserAgent[UA_VERSION],MAX_UA_LEN,Value);
				bCustomUserAgent[UA_VERSION]=true;
				IsCustomUserAgentSpecified=true;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("ua-platform")))
			{
				StringCchCopy(CustomUserAgent[UA_PLATFORM],MAX_UA_LEN,Value);
				bCustomUserAgent[UA_PLATFORM]=true;
				IsCustomUserAgentSpecified=true;
				IsGlobalTaskPresent=true;
			}
			else if(!_tcsicmp(Option,_T("ua-reset")))
			{
				for(int i=0;i<UA_NFIELDS;i++)
				{
					CustomUserAgent[i][0]=0;
					bCustomUserAgent[i]=true;
				}
				IsCustomUserAgentSpecified=true;
				IsGlobalTaskPresent=true;
			}
			else
			{
				// Assume that argument is an address or file
				if(IsFirstProperty && !Value[0] && (!ch || ch==L' '))
				{
					InstanceProperties *pInstance;

					// Parse URL from internet shortcut
					if(_tcsicmp(PathFindExtension(Option),_T(".url")))
					{
						ParseURLFile(Option);
					}

					// Append argument (Option) to all instances
					i=0;
					pInstance=Instances;
					while(i<=n)
					{
						if(pInstance->Options&OPTION_STARTPROCESS)
						{
							// Append argument to command
							StringCbCat(pInstance->Open,sizeof pInstance->Open,_T(" "));
							StringCbCat(pInstance->Open,sizeof pInstance->Open,Option);
						}
						else
						{
							StripQutoes(pInstance->Open,Option);
							pInstance->bOpen=true;
						}
						
						i++;
						pInstance++;
					}
					if(n==0) IsInstance=true;
				}

			}

			// Skip separator
			while(ch==L',')
			{
				pCmdLine++;
				ch=*pCmdLine;
			}

			IsFirstProperty=false;

		} while(ch && ch!=L' ');

		while(_istspace(*pCmdLine)) pCmdLine++;
		
		if(IsInstance)
		{
			n++;
			pCurrentInstance++;
		}
	}

	// Delete Traces
	if(DeleteTracesOptions) DeleteTraces(DeleteTracesOptions);

	// Set Custom User Agent
	if(IsCustomUserAgentSpecified) WriteUserAgent(CustomUserAgent,bCustomUserAgent);

	if(n)
	{
		// Start IE instances
		i=n-1;
		pCurrentInstance=Instances+i;
		while(i>=0)
		{
			// Start instance that will receive focus last to mitigate focus stealing
			if(i==0) pCurrentInstance=Instances+Focus;
			else if(i==Focus) pCurrentInstance--;

			if((pCurrentInstance->Options&OPTION_STARTPROCESS)==0)
			{
				hr=CoCreateInstance(CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, (void**)&pIE);
				
				if(SUCCEEDED(hr))
				{
					VARIANT_BOOL vBool;
					int iBool;

					pIE->get_HWND((LONG*)&hwnd);				

					// IE7 get parent window

					parenthwnd=GetParent(hwnd);
					if(parenthwnd) hwnd=parenthwnd;

					if(pCurrentInstance->Options&OPTION_RESIZEABLE_SET)
					{
						if(pCurrentInstance->Options&OPTION_RESIZEABLE) vBool=VARIANT_TRUE;
						else vBool=VARIANT_FALSE;

						pIE->put_Resizable(vBool);
					}

					PrepareWindow(hwnd,pCurrentInstance);

					if(pCurrentInstance->Options&OPTION_QTOOLBAR_SET)
					{
						VARIANT vCLSID,vShow,vSize;

						if(pCurrentInstance->Options&OPTION_QTOOLBAR) vBool=VARIANT_TRUE;
						else vBool=VARIANT_FALSE;

						vCLSID.vt=VT_BSTR;
						vCLSID.bstrVal=CComBSTR(L"{A411D7F4-8D11-43EF-BDE4-AA921666388A}");
						vShow.vt=VT_BOOL;
						vShow.boolVal=vBool;
						vSize.vt=VT_EMPTY;

						pIE->ShowBrowserBar(&vCLSID,&vShow,&vSize);
					}
				
					if(pCurrentInstance->Options&OPTION_TOOLBAR_SET)
					{
						if(pCurrentInstance->Options&OPTION_TOOLBAR) iBool=TRUE;
						else iBool=FALSE;

						pIE->put_ToolBar(iBool);
					}

					if(pCurrentInstance->Options&OPTION_LOCATION_SET)
					{
						if(pCurrentInstance->Options&OPTION_LOCATION) vBool=VARIANT_TRUE;
						else vBool=VARIANT_FALSE;

						pIE->put_AddressBar(vBool);
					}

					if(pCurrentInstance->Options&OPTION_MENUBAR_SET)
					{
						if(pCurrentInstance->Options&OPTION_MENUBAR) vBool=VARIANT_TRUE;
						else vBool=VARIANT_FALSE;

						pIE->put_MenuBar(vBool);
					}

					if(pCurrentInstance->Options&OPTION_STATUS_SET)
					{
						if(pCurrentInstance->Options&OPTION_STATUS) vBool=VARIANT_TRUE;
						else vBool=VARIANT_FALSE;

						pIE->put_StatusBar(vBool);
					}

					if(pCurrentInstance->Options&OPTION_EXPLORERBAR_SET)
					{
						VARIANT vCLSID,vShow,vSize;

						vCLSID.vt=VT_BSTR;
						vShow.vt=VT_BOOL;
						vSize.vt=VT_EMPTY;
						if(pCurrentInstance->ExplorerBar==EXPLORERBAR_NONE)
						{
							int j;

							vShow.boolVal=VARIANT_FALSE;

							for(j=1;j<MAX_EXPLORERBARS;j++)
							{
								vCLSID.bstrVal=CComBSTR(ExplorerBars[j].CLSID);
								pIE->ShowBrowserBar(&vCLSID,&vShow,&vSize);
							}
						}
						else if(pCurrentInstance->ExplorerBar<MAX_EXPLORERBARS)
						{
							vShow.boolVal=VARIANT_TRUE;
							vCLSID.bstrVal=CComBSTR(ExplorerBars[pCurrentInstance->ExplorerBar].CLSID);
							pIE->ShowBrowserBar(&vCLSID,&vShow,&vSize);
						}
					}

					if(pCurrentInstance->Options&OPTION_SILENT_SET)
					{
						if(pCurrentInstance->Options&OPTION_SILENT) vBool=VARIANT_TRUE;
						else vBool=VARIANT_FALSE;

						pIE->put_Silent(vBool);
					}

					if(pCurrentInstance->Options&OPTION_KIOSK) pIE->put_FullScreen(VARIANT_TRUE);
					if(pCurrentInstance->Options&OPTION_THEATER) pIE->put_TheaterMode(VARIANT_TRUE);
	
					if(pCurrentInstance->bOpen)
					{
						VARIANT vEmpty,vFlags;
						TCHAR *pOpen;
						TCHAR ch;
						int j,k;
						TCHAR URL[MAX_URL_LEN];

						VariantInit(&vEmpty);
						vFlags.vt=VT_I4;

						pOpen=pCurrentInstance->Open;
						j=0;
						k=0;

						// Split string into URLs
						do
						{
							ch=*pOpen;

							if(ch==_T('|') || ch==_T('\0'))
							{
								if(k)
								{
									URL[k]=0;

									// Open tab
									vFlags.intVal=j?navOpenBackgroundTab:0;
									pIE->Navigate(CComBSTR(URL), &vFlags, &vEmpty, &vEmpty, &vEmpty);

									k=0;
									j++;
								}
							}
							else
							{
								if(k<MAX_URL_LEN-1)	URL[k++]=ch;
								else break; // URL too long
							}

							pOpen++;
						} while(ch && j<MAX_IE_TABS);
					}
					else pIE->GoHome();

					pCurrentInstance->hWnd=hwnd;
				}
				else
				{
					MessageBeep(MB_ICONEXCLAMATION);
					pCurrentInstance->hWnd=NULL;
					ExitCode=EXIT_ERROR;
				}
			}
			else pCurrentInstance->hWnd=NULL;
			i--;
			pCurrentInstance--;
		}


		// Make the opened IE instances visible

		i=0;
		pCurrentInstance=Instances;
		while(i<n)
		{
			if(pCurrentInstance->hWnd)
			{
				ShowWindow(pCurrentInstance->hWnd,pCurrentInstance->ShowCommand);

				// Remove IE7 Navigation Bar
				/*

				HWND IE7NavigationBarHwnd;

					IE7NavigationBarHwnd=FindWindowEx(pCurrentInstance->hWnd,NULL,L"WorkerW",L"Navigation Bar");
					if(IE7NavigationBarHwnd)
					{
						//MessageBox(NULL,L"navbar found",L"quero",MB_OK);
						ShowWindow(IE7NavigationBarHwnd,SW_HIDE);
						//MoveWindow(IE7NavigationBarHwnd,0,0,0,0,TRUE);

						HWND hWndReBar=FindWindowEx(IE7NavigationBarHwnd,NULL,L"ReBarWindow32",NULL);

						//if(hWndReBar) MessageBox(NULL,L"rebar found",L"quero",MB_OK);

						int nCount = (int)::SendMessage(hWndReBar, RB_GETBANDCOUNT, 0, 0L);
						for(int i = 0; i < nCount; i++)
						{
							REBARBANDINFO rbbi = { sizeof(REBARBANDINFO), RBBIM_CHILD | RBBIM_CHILDSIZE | RBBIM_STYLE};
							BOOL bRet = (BOOL)::SendMessage(hWndReBar, RB_GETBANDINFO, i, (LPARAM)&rbbi);

							//MessageBox(NULL,L"",L"rb_getbandinfo",MB_OK);
							
							// if(bRet) fails?
							{
								//MessageBox(NULL,L"",L"setbandinfo",MB_OK);
								rbbi.fMask=RBBIM_CHILDSIZE | RBBIM_STYLE;
								rbbi.fStyle= RBBS_HIDDEN | RBBS_VARIABLEHEIGHT | RBBS_GRIPPERALWAYS;
									rbbi.cyMinChild = rbbi.cyChild = 64; 
     								//::SendMessage(hWndReBar, RB_SETBANDINFO, i, (LPARAM)&rbbi);

									::SendMessage(hWndReBar, RB_SHOWBAND, i, 0);
							}
						}

						RECT rect;
						HWND hwnd;

						hwnd=GetParent(IE7NavigationBarHwnd);

						GetWindowRect(hwnd,&rect);
						SendMessage(hwnd,WM_SIZE,0,MAKELPARAM(rect.bottom-rect.top,rect.right-rect.left));

						ShowWindow(IE7NavigationBarHwnd,SW_SHOW);


						
					}

				*/
			}
			i++;
			pCurrentInstance++;
		}

		// Start other processes

		if(nStartProcess)
		{
			i=0;
			pCurrentInstance=Instances;
			while(i<n)
			{
				if(pCurrentInstance->Options&OPTION_STARTPROCESS)
				{
					STARTUPINFO si; // si and pi structures for CraeteProcess
					PROCESS_INFORMATION pi;

					ZeroMemory( &si, sizeof(si) );
					si.cb = sizeof(si);

					si.wShowWindow=pCurrentInstance->ShowCommand;
					si.dwFlags=STARTF_USESHOWWINDOW;

					if(CreateProcess(NULL,pCurrentInstance->Open,NULL,NULL,FALSE,0,NULL,NULL,&si,&pi)==FALSE)
					{
						MessageBox(NULL,pCurrentInstance->Open,_T("Quero Launcher: failed to start process"),MB_OK|MB_ICONEXCLAMATION);
						ExitCode=EXIT_ERROR;
					}
					else
					{
						pCurrentInstance->ProcessInfo=pi;
					}
				}
				i++;
				pCurrentInstance++;
			}

			// Grab window handles
		
			int retry;

			retry=0;

			while(retry<MAX_RETRIES)
			{
				bool TryAgain;

				TryAgain=false;
				i=0;
				pCurrentInstance=Instances;
				while(i<n)
				{
					if(pCurrentInstance->Options&OPTION_STARTPROCESS &&
						(pCurrentInstance->bLeft||pCurrentInstance->bTop||pCurrentInstance->bWidth||pCurrentInstance->bHeight||pCurrentInstance->Options!=OPTION_DEFAULT))
					{
						if(pCurrentInstance->ProcessInfo.dwThreadId)
						{
							DWORD ThreadExitCode;

							// Has Thread terminated?
							if(GetExitCodeThread(pCurrentInstance->ProcessInfo.hThread,&ThreadExitCode)==0 || ThreadExitCode!=STILL_ACTIVE)
							{
								pCurrentInstance->ProcessInfo.dwThreadId=0;
							}
							else
							{
								EnumThreadWindows(pCurrentInstance->ProcessInfo.dwThreadId,FoundThreadHwnd,(LPARAM)&pCurrentInstance->hWnd);
								if(pCurrentInstance->hWnd)
								{
									//MessageBox(NULL,pCurrentInstance->Open,L"found",MB_OK);
									PrepareWindow(pCurrentInstance->hWnd,pCurrentInstance);

									ShowWindow(pCurrentInstance->hWnd,pCurrentInstance->ShowCommand);

									pCurrentInstance->ProcessInfo.dwThreadId=0;
								}
								else TryAgain=true;
							}
						}
					}

					i++;
					pCurrentInstance++;
				}

				if(TryAgain==false) break;

				Sleep(WAITFORSTARTUP);
				retry++;
			}

			// Close the process handles
			i=0;
			pCurrentInstance=Instances;
			while(i<n)
			{
				if(pCurrentInstance->Options&OPTION_STARTPROCESS)
				{
					CloseHandle(pCurrentInstance->ProcessInfo.hProcess);
					CloseHandle(pCurrentInstance->ProcessInfo.hThread);
				}
				i++;
				pCurrentInstance++;
			}
		}
		
		// Activate window
		hwnd=NULL;
		if(bFocus) hwnd=Instances[Focus].hWnd;
		else if(nStartProcess==0)
		{
			i=0;
			pCurrentInstance=Instances;
			while(i<n && (pCurrentInstance->hWnd==NULL || pCurrentInstance->ShowCommand==SW_MINIMIZE || pCurrentInstance->ShowCommand==SW_SHOWMINNOACTIVE))
			{
				i++;
				pCurrentInstance++;
			}
			if(i<n) hwnd=pCurrentInstance->hWnd;
		}
		if(hwnd) SetForegroundWindow(hwnd);
	}
	else if(IsGlobalTaskPresent==false)
		MessageBox(NULL,_T("Launches Internet Explorer and other applications.\n\n")
		_T("Create a shortcut to Quero Launcher and edit the Target field to specify what to launch.\n")
		_T("\n")
		_T("Usage: QLaunch [General_Options] [Options_Window1] [Options_Window2] ...\n")
		_T("\n")
		_T("Options\n")
		_T("\t[property=value | property],...\n\tcomma-separated list of properties\n")
		_T("\n")
		_T("General\n")
		_T("\tdeletehistory | deletecache | deletecookies | deletepasswords | deleteforms | deletefavorites | deleteall |\n")
		_T("\tua-prefix=string | ua-compatible=string | ua-version=string | ua-platform=string | ua-reset\n\n")
		_T("Properties\n")
		_T("\topen[=URLs separated by '|' character] | start=command_line_with_fully_qualified_path |\n")
		_T("\tleft=value | top=value | width=value | height=value |\n")
		_T("\tfocus | maximized | minimized | alwaysontop | notitlebar | resizeable=yes|no\n")
		_T("\n")
		_T("IE-specific Window Properties\n")
		_T("\tlocation=yes|no | menubar=yes|no | silent=yes|no | status=yes|no | toolbar=yes|no |\n")
		_T("\tqtoolbar=yes|no | explorerbar=Search|Favorites|History|Channels|Folders|none | kiosk | theater\n")
		_T("\n")
		_T("IE7 Customization\n")
		_T("\ttabbedbrowsing=yes|no | searchbox=yes|no | commandbar=yes|no |\n")
		_T("\titbar7position=value (0 default, 1 place toolbars above the navigation bar in IE7)\n")
		_T("\n")
		_T("Examples\n\tLaunch IE maximized\n\tQLaunch maximized\n\n")
		_T("\tLaunch two IE windows side by side\n\tQLaunch left=0,top=0,width=50%,height=100% left=50%,top=0,width=50%,height=100%\n\n")
		_T("\tOpen three tabs in IE7\n\tQLaunch open=\"microsoft.com|xbox.com|live.com\"\n\n")
#ifdef COMPILE_FOR_WIN9X
		_T("Version 1.0.6 Win9x Build --- http://www.quero.at/ --- Copyright 2006 Viktor Krammer"),_T("Quero Launcher"),MB_ICONINFORMATION|MB_OK);
#else
		_T("Version 1.0.6 --- http://www.quero.at/ --- Copyright 2006 Viktor Krammer"),_T("Quero Launcher"),MB_ICONINFORMATION|MB_OK);
#endif

	CoUninitialize();

	return ExitCode;
}

void PrepareWindow(HWND hwnd,InstancePropertie *pInstance)
{
	RECT rect;
	int x,y,w,h;
	HWND ZOrder;
	UINT uFlags;
	LONG style;

	GetWindowRect(hwnd,&rect);
	x=rect.left;
	y=rect.top;
	w=rect.right-x;
	h=rect.bottom-y;

	if(pInstance->bLeft) x=pInstance->Left;
	if(pInstance->bTop) y=pInstance->Top;
	if(pInstance->bWidth) w=pInstance->Width;
	if(pInstance->bHeight) h=pInstance->Height;

	if(pInstance->Options&OPTION_ALWAYSONTOP) ZOrder=HWND_TOPMOST;
	else ZOrder=HWND_TOP;

	uFlags=SWP_NOACTIVATE;

	if(pInstance->Options&OPTION_NOTITLEBAR)
	{
		style=GetWindowLong(hwnd,GWL_STYLE);
		style&=~(WS_CAPTION); // Remove the titlebar
		SetWindowLong(hwnd,GWL_STYLE,style);

		uFlags|=SWP_FRAMECHANGED|SWP_DRAWFRAME;
	}

	if(pInstance->Options&(OPTION_RESIZEABLE_SET))
	{
		uFlags|=SWP_FRAMECHANGED|SWP_DRAWFRAME;

		if(pInstance->Options&OPTION_STARTPROCESS)
		{
			style=GetWindowLong(hwnd,GWL_STYLE);
			if(pInstance->Options&OPTION_RESIZEABLE)
			{
				style|=WS_THICKFRAME|WS_MAXIMIZEBOX;
			}
			else
			{
				style&=~(WS_THICKFRAME|WS_MAXIMIZEBOX); // Remove the titlebar
				style|=WS_BORDER;
			}
			SetWindowLong(hwnd,GWL_STYLE,style);
		}
	}

	SetWindowPos(hwnd,ZOrder,x,y,w,h,uFlags);
}


BOOL CALLBACK FoundThreadHwnd(HWND hwnd,LPARAM l)
{
	LONG exstyle;

	if(GetParent(hwnd)==NULL)
	{
		exstyle=GetWindowLong(hwnd,GWL_EXSTYLE);

		// Is the window visible and shown in the taskbar?
		if(IsWindowVisible(hwnd) && ((exstyle&WS_EX_TOOLWINDOW)==0 || (exstyle&WS_EX_APPWINDOW)))
		{
			// Check WindowTextLength?		
			*(HWND*)l=hwnd;
			return false; // Stop enumeration
		}
	}
	return true;
}

void StripQutoes(TCHAR *pDst,TCHAR *pSrc)
{
	TCHAR last,current;
	int j;

	if(*pSrc==L'"') pSrc++;
	j=0;
	current=last=*pSrc;
	while(current && j<MAX_URL_LEN-1)
	{
		*pDst=current;
		pDst++;
		last=current;
		pSrc++;
		current=*pSrc;
		j++;
	}
	if(last==L'"') pDst--;
	*pDst=L'\0';
}

void ParseURLFile(TCHAR *pOption)
{
	HANDLE hFile;
	DWORD BytesRead;
	char *pBuffer;
	char KeyWord[]="URL=";
	int iKeyWord;
	int iURL;
	int state;
	char ch;
	char buffer[MAX_URL_LEN*3];
	TCHAR URL[MAX_URL_LEN];
	TCHAR FileName[MAX_URL_LEN];

	StripQutoes(FileName,pOption);
	hFile=CreateFile(FileName,GENERIC_READ,FILE_SHARE_READ,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
	if(hFile!=INVALID_HANDLE_VALUE)
	{
		if(ReadFile(hFile,buffer,sizeof buffer,&BytesRead,NULL))
		{
			state=1;
			iURL=0;
			pBuffer=buffer;
			while(BytesRead)
			{
				ch=*pBuffer;
				switch(state)
				{
				case 0:
					if(ch=='\n')
					{
						iKeyWord=0;
						state=1;
					}
					break;
				case 1:
					if(ch==KeyWord[iKeyWord])
					{
						iKeyWord++;
						if(iKeyWord+1==sizeof KeyWord) state=3;
					}
					else state=0;
					break;
				case 3:
					if(ch==L'\r' || ch==L'\n' || iURL>=MAX_URL_LEN)
					{
						URL[iURL]=L'\0';
						BytesRead=1; // Finished
					}
					URL[iURL]=ch;
					iURL++;
					break;
				}

				pBuffer++;
				BytesRead--;
			}

			if(iURL)
			{
				StringCchCopy(pOption,MAX_URL_LEN,URL);
			}
		}
		CloseHandle(hFile);
	}
}

void DeleteTraces(UINT clear)
{
	HKEY hKeyLocal;
	HRESULT hr;

	if(clear&CLEAR_HISTORY)
	{
		RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software"), 0, NULL, REG_OPTION_NON_VOLATILE, DELETE, NULL, &hKeyLocal, NULL);
		RegDeleteKey(hKeyLocal,_T("Quero Toolbar\\History"));
		RegDeleteKey(hKeyLocal,_T("Quero Toolbar\\URL"));
		RegDeleteKey(hKeyLocal,_T("Microsoft\\Internet Explorer\\TypedURLs"));
		RegCloseKey(hKeyLocal);

		IUrlHistoryStg2* pUrlHistoryStg2;

		pUrlHistoryStg2=NULL;
		hr = CoCreateInstance(CLSID_CUrlHistory,NULL,CLSCTX_INPROC_SERVER, IID_IUrlHistoryStg2,(void**)&pUrlHistoryStg2);
		if(SUCCEEDED(hr))
		{
			hr = pUrlHistoryStg2->ClearHistory();
			pUrlHistoryStg2->Release();
		}
	}
	if(clear&(CLEAR_CACHE|CLEAR_COOKIES))
	{
		DeleteUrlCache(clear);
	}
	if(clear&(CLEAR_AUTOCOMPLETE_FORMS|CLEAR_AUTOCOMPLETE_PWDS))
	{
		IWebBrowser2 *pIE;

		hr=CoCreateInstance(CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, (void**)&pIE);
		if(SUCCEEDED(hr))
		{
			IOleCommandTarget *pOleCommandTarget;
			VARIANTARG vaIn;

			vaIn.vt=VT_I4;
			switch(clear&(CLEAR_AUTOCOMPLETE_FORMS|CLEAR_AUTOCOMPLETE_PWDS))
			{
			case CLEAR_AUTOCOMPLETE_FORMS:
				vaIn.intVal=IECMDID_ARG_CLEAR_FORMS_ALL_BUT_PASSWORDS;
				break;
			case CLEAR_AUTOCOMPLETE_PWDS:
				vaIn.intVal=IECMDID_ARG_CLEAR_FORMS_PASSWORDS_ONLY;
				break;
			default:
				vaIn.intVal=IECMDID_ARG_CLEAR_FORMS_ALL;
			}

			hr=pIE->QueryInterface(IID_IOleCommandTarget,(LPVOID*)&pOleCommandTarget);
			if(SUCCEEDED(hr))
			{
				pOleCommandTarget->Exec(&CGID_InternetExplorer,IECMDID_CLEAR_AUTOCOMPLETE_FOR_FORMS,OLECMDEXECOPT_DONTPROMPTUSER,&vaIn,NULL);
				pOleCommandTarget->Release();
			}
 		
			pIE->Quit();
		}
	}
	if(clear&(CLEAR_FAVORITES))
	{
		TCHAR path[MAX_PATH];

#ifdef COMPILE_FOR_WIN9X
		if(SHGetSpecialFolderPath(NULL,path,CSIDL_FAVORITES,FALSE)==TRUE)
		{		
			EmptyDirectory(path,false);
		}
#else
		hr=SHGetFolderPath(NULL,CSIDL_FAVORITES,NULL,SHGFP_TYPE_CURRENT,path);
		if(SUCCEEDED(hr))
		{
			EmptyDirectory(path,false);
		}
#endif
	}
}

bool DeleteUrlCache(UINT type)
{
	HANDLE hCacheEnum;
	TCHAR CacheEntry[MAX_INET_CACHE_ENTRY_SIZE];
    LPINTERNET_CACHE_ENTRY_INFO pCacheEntry;
    DWORD dwEntrySize;
	bool ret=false;
#ifndef COMPILE_FOR_WIN9X
	HRESULT hr;
#endif
	TCHAR path[MAX_PATH];

	pCacheEntry=(LPINTERNET_CACHE_ENTRY_INFO)&CacheEntry;
	dwEntrySize=sizeof CacheEntry;
	hCacheEnum=FindFirstUrlCacheEntry(NULL,pCacheEntry,&dwEntrySize);
	if(hCacheEnum)
	{
		do
		{
			if(pCacheEntry->CacheEntryType&COOKIE_CACHE_ENTRY)
			{
				if(type&CLEAR_COOKIES) DeleteUrlCacheEntry(pCacheEntry->lpszSourceUrlName);
			}
			else
			{
				if(type&CLEAR_CACHE) DeleteUrlCacheEntry(pCacheEntry->lpszSourceUrlName);
			}

			dwEntrySize=sizeof CacheEntry;
		}
		while(FindNextUrlCacheEntry(hCacheEnum,pCacheEntry,&dwEntrySize));
		if(GetLastError()==ERROR_NO_MORE_ITEMS) ret=true;

		FindCloseUrlCache(hCacheEnum);
	}

	if(type&CLEAR_COOKIES)
	{
#ifdef COMPILE_FOR_WIN9X
		if(SHGetSpecialFolderPath(NULL,path,CSIDL_COOKIES,FALSE)==TRUE)
		{
			EmptyDirectory(path,false);
		}
#else
		hr=SHGetFolderPath(NULL,CSIDL_COOKIES,NULL,SHGFP_TYPE_CURRENT,path);
		if(SUCCEEDED(hr))
		{
			EmptyDirectory(path,false);
		}
#endif
	}
	if(type&CLEAR_CACHE)
	{
#ifdef COMPILE_FOR_WIN9X
		if(SHGetSpecialFolderPath(NULL,path,CSIDL_INTERNET_CACHE,FALSE)==TRUE)
		{
			EmptyDirectory(path,false);
		}
#else
		hr=SHGetFolderPath(NULL,CSIDL_INTERNET_CACHE,NULL,SHGFP_TYPE_CURRENT,path);
		if(SUCCEEDED(hr))
		{
			EmptyDirectory(path,false);
		}
#endif
	}

	return ret;
}

bool EmptyDirectory(LPCTSTR szPath,bool bDeleteDesktopIni)
{
	WIN32_FIND_DATA wfd;
	HANDLE hFind;
	TCHAR FindFilter[MAX_PATH];
	TCHAR FullPath[MAX_PATH];
	DWORD dwAttributes = 0;

	StringCbCopy(FindFilter,sizeof FindFilter,szPath);
	PathAppend(FindFilter,_T("*.*"));

	if((hFind = FindFirstFile(FindFilter, &wfd)) != INVALID_HANDLE_VALUE)
	{
		do
		{
			if(_tcscmp(wfd.cFileName,_T(".")) && _tcscmp(wfd.cFileName,_T("..")) &&
				(bDeleteDesktopIni || _tcsicmp(wfd.cFileName,_T("desktop.ini"))))
			{
				StringCbCopy(FullPath,sizeof FullPath,szPath);
				PathAppend(FullPath,wfd.cFileName);
			
				//Remove readonly attribute
				dwAttributes = GetFileAttributes(FullPath);
				if(dwAttributes & FILE_ATTRIBUTE_READONLY)
				{
					dwAttributes &= ~FILE_ATTRIBUTE_READONLY;
					SetFileAttributes(FullPath, dwAttributes);
				}

				if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // Recursivly delete subdirs
				{
					EmptyDirectory(FullPath, bDeleteDesktopIni);
					RemoveDirectory(FullPath);
				}
				else
				{
					DeleteFile(FullPath);
				}
			}
		}
		while(FindNextFile(hFind, &wfd));

		FindClose(hFind);
	}
	else return false;

	return true;
}


void WriteUserAgent(TCHAR CustomUserAgent[UA_NFIELDS][MAX_UA_LEN],bool bCustomUserAgent[UA_NFIELDS])
{
	TCHAR* UA_Values[UA_NFIELDS]={NULL,_T("Compatible"),_T("Version"),_T("Platform")};
	HKEY hKeyUserAgent;
	int i;

	if(SUCCEEDED(RegCreateKeyEx(HKEY_LOCAL_MACHINE,_T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\5.0\\User Agent"),0,NULL,REG_OPTION_NON_VOLATILE,KEY_WRITE,NULL,&hKeyUserAgent,NULL)))
	{
		for(i=0;i<UA_NFIELDS;i++)
		{
			if(bCustomUserAgent[i])
			{
				if(CustomUserAgent[i][0]==0) RegDeleteValue(hKeyUserAgent,UA_Values[i]);
				else RegSetValueEx(hKeyUserAgent,UA_Values[i],0,REG_SZ,(LPBYTE)CustomUserAgent[i],(_tcslen(CustomUserAgent[i])+1)*sizeof TCHAR);
			}
		}

		RegCloseKey(hKeyUserAgent);
	}
}