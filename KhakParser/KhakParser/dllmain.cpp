// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "KhakParser_i.h"
#include "dllmain.h"
#include "CKhParser.h"

CKhakParserModule _AtlModule;

OBJECT_ENTRY_AUTO(CLSID_KhParser, CKhParser )

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved); 
}
