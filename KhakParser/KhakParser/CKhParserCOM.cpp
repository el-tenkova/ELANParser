#include "stdafx.h"

#include "KhakParser_i.h"
#include "CKhParserCOM.h"

STDMETHODIMP CKhParserCOM::Init(int cSafeArr, BSTR www, BSTR dictPath, BSTR notfoundPath, long* hRes)
{
    safeArraySize = cSafeArr;
    *hRes = pureParser.Init(www, dictPath, notfoundPath);
    safeArraySize = cSafeArr;
    return *hRes;
}

STDMETHODIMP CKhParserCOM::Terminate(long* hRes)
{
    *hRes = pureParser.Terminate();
    return S_OK;
}

HRESULT CKhParserCOM::DoParse(BSTR word, long* hRes)
{
    *hRes = pureParser.DoParse(word);
     return S_OK;
}

HRESULT CKhParserCOM::GetHomonymNum(int* cNumHom)
{
    *cNumHom = pureParser.GetHomonymNum();
    return S_OK;
}

HRESULT CKhParserCOM::GetNextHomonym(SAFEARRAY** lpHomonym)
{
    std::wstring str;
    pureParser.GetNextHomonym(str);

    long idx, len, start;
    len = (long)str.length();

    SafeArrayLock(*lpHomonym);

    SafeArrayGetUBound(*lpHomonym, 1, &len);
    SafeArrayGetLBound(*lpHomonym, 1, &start);
    short asterisk = L'*';
    for (idx = start; idx < len; idx++) {
        if (idx == str.length()) {
            //        if (str[idx] == 0x0) {
            SafeArrayPutElement(*lpHomonym, &idx, &asterisk);
            break;
        }
        SafeArrayPutElement(*lpHomonym, &idx, &str[idx]);
    }
    SafeArrayUnlock(*lpHomonym);
    return S_OK;
}

HRESULT CKhParserCOM::Normalize(BSTR InputWord, SAFEARRAY** lpHomonym)
{
    return S_OK;
}

HRESULT CKhParserCOM::AddKhakSent(BSTR InputSent, long* hRes)
{
    *hRes = pureParser.AddKhakSent(InputSent);
    return S_OK;
}

HRESULT CKhParserCOM::AddKhakSent2(BSTR Name, BSTR InputSent, long* hRes)
{
    *hRes = pureParser.AddKhakSent2(Name, InputSent);
    return S_OK;
}

HRESULT CKhParserCOM::AddRusSent(BSTR InputSent, long* hRes)
{
    *hRes = pureParser.AddRusSent(InputSent);
    return S_OK;
}
HRESULT CKhParserCOM::SaveToELAN(BSTR ElanPath, /*[out, retval]*/ long *hRes)
{
    *hRes = pureParser.SaveToELAN(ElanPath);
    return S_OK;
}