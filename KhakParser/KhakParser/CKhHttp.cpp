#include "stdafx.h"

#ifdef _USE_MSXML2
#include "CKhHttp.h"

CKhHttpWrapper::CKhHttpWrapper(long* hRes, const char* _www)
{
    long res = pIXMLHTTPRequest.CreateInstance("Msxml2.ServerXMLHTTP");
    if (hRes != 0)
        *hRes = res;
    www.append(_www);
}
CKhHttpWrapper::~CKhHttpWrapper(void)
{
    if (pIXMLHTTPRequest != nullptr)
        pIXMLHTTPRequest.Release();
}
long CKhHttpWrapper::Open(const wchar_t* toPost)
{
    return pIXMLHTTPRequest->open("GET", toPost, false);
}
long CKhHttpWrapper::Send(void)
{
    return pIXMLHTTPRequest->send();
}
size_t CKhHttpWrapper::Status(void)
{
    return pIXMLHTTPRequest->status;
}
const wchar_t* CKhHttpWrapper::ResponseText(void)
{
    return pIXMLHTTPRequest->responseText;
}
#endif
