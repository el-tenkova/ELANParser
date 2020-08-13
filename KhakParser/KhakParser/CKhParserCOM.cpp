#include "stdafx.h"

#include "KhakParser_i.h"
#include "CKhParserCOM.h"

STDMETHODIMP CKhParserCOM::Init(int cSafeArr, BSTR _www, BSTR _dictPath, BSTR _notfoundPath, long* hRes)
{
    safeArraySize = cSafeArr;

    symbols.insert(std::pair<wchar_t, wchar_t>(0x0492, 0x0413));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x0493, 0x0433));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04CB, 0x0427));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04CC, 0x0447));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04A3, 0x043D));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04A2, 0x041D));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04E7, 0x043E));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04E6, 0x041E));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04F1, 0x0443));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x04F0, 0x0423));
    symbols.insert(std::pair<wchar_t, wchar_t>(0x0456, 0x0438));

    mapFileName(_dictPath);
    mapFileName(_notfoundPath);
    char tmp[512] = "";
    size_t len;
    _locale_t locale = _create_locale(LC_ALL, "ru-RU");
    errno_t err = _wcstombs_s_l(&len, tmp, sizeof(tmp), _dictPath, _TRUNCATE, locale);
    std::string dictPath(tmp);
    *tmp = 0x0;
    len = 0;
    _wcstombs_s_l(&len, tmp, sizeof(tmp), _notfoundPath, _TRUNCATE, locale);
    std::string notfoundPath(tmp);
    *hRes = pureParser.Init(_www, dictPath, notfoundPath);
    safeArraySize = cSafeArr;
    _free_locale(locale);
    return *hRes;
}

STDMETHODIMP CKhParserCOM::Terminate(long* hRes)
{
    *hRes = pureParser.Terminate();
    return S_OK;
}

HRESULT CKhParserCOM::DoParse(BSTR word, long* hRes)
{
    if (word == 0)
        return S_OK;
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

HRESULT CKhParserCOM::AddKhakSent3(BSTR Name, BSTR Time, BSTR InputSent, long* hRes)
{
    *hRes = pureParser.AddKhakSent3(Name, Time, InputSent);
    return S_OK;
}

HRESULT CKhParserCOM::AddRusSent(BSTR InputSent, long* hRes)
{
    *hRes = pureParser.AddRusSent(InputSent);
    return S_OK;
}

HRESULT CKhParserCOM::AddTranscription(BSTR InputSent, long* hRes)
{
    *hRes = pureParser.AddTranscription(InputSent);
    return S_OK;
}

HRESULT CKhParserCOM::SaveToELAN(BSTR ElanPath, /*[out, retval]*/ long *hRes)
{
    mapFileName(ElanPath);
    _locale_t locale = _create_locale(LC_ALL, "ru-RU");
    char tmp[512] = "";
    size_t len = 0;
    _wcstombs_s_l(&len, tmp, sizeof(tmp), ElanPath, _TRUNCATE, locale);
    std::string elanPath(tmp);
    *hRes = pureParser.SaveToELAN(elanPath);
    _free_locale(locale);
    return S_OK;
}
HRESULT CKhParserCOM::SaveToELANFlex(BSTR ElanPath, /*[out, retval]*/ long *hRes)
{
    mapFileName(ElanPath);
    _locale_t locale = _create_locale(LC_ALL, "ru-RU");
    char tmp[512] = "";
    size_t len = 0;
    _wcstombs_s_l(&len, tmp, sizeof(tmp), ElanPath, _TRUNCATE, locale);
    std::string elanPath(tmp);
    *hRes = pureParser.SaveToELANFlex(elanPath);
    _free_locale(locale);
    return S_OK;
}
HRESULT CKhParserCOM::SaveToELANFlexTime(BSTR ElanPath, /*[out, retval]*/ long *hRes)
{
    mapFileName(ElanPath);
    _locale_t locale = _create_locale(LC_ALL, "ru-RU");
    char tmp[512] = "";
    size_t len = 0;
    errno_t err = _wcstombs_s_l(&len, tmp, sizeof(tmp), ElanPath, _TRUNCATE, locale);
    err++;
    std::string elanPath(tmp);
    *hRes = pureParser.SaveToELANFlexTime(elanPath);
    _free_locale(locale);
    return S_OK;
}
void CKhParserCOM::mapFileName(BSTR fileName)
{
    if (fileName == 0)
        return;
    size_t len = wcslen(fileName);
    for (size_t i = 0; i < len; i++)
    {
        std::map<wchar_t, wchar_t>::iterator it = symbols.find(fileName[i]);
        if (it != symbols.end())
            fileName[i] = it->second;
    }
}
