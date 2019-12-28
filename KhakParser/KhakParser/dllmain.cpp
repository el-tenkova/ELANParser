// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "KhakParser_i.h"
#include "dllmain.h"
#include "CKhParserCOM.h"
//#include "CLingvo.h"

CKhakParserModule _AtlModule;

OBJECT_ENTRY_AUTO(CLSID_KhParser, CKhParserCOM )
//OBJECT_ENTRY_AUTO(CLSID_Lingvo, CLingvo)
// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved); 
}
