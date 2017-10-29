// CKhParser.cpp : Implementation of CKhParser

#include "stdafx.h"
#include <locale.h>
#include <fstream>
#include <codecvt>

#include "KhakParser_i.h"
#include "CKhParser.h"

STDMETHODIMP CKhParser::Init(int cSafeArr, BSTR www, BSTR dictPath, BSTR notfoundPath, long* hRes )
{
    Terminate( hRes );
    dict = std::wstring(dictPath);
    notfound = std::wstring(notfoundPath);
    //"http://khakas.altaica.ru"
    request = www + request;
    safeArraySize = cSafeArr;
    *hRes = pIXMLHTTPRequest.CreateInstance("Msxml2.ServerXMLHTTP");
    
    repl.insert(std::pair<short, short>(L'a', 0x0430));
    repl.insert(std::pair<short, short>(L'c', 0x0441));
    repl.insert(std::pair<short, short>(L'e', 0x0435));
    repl.insert(std::pair<short, short>(L'i', 0x0456));
    repl.insert(std::pair<short, short>(L'o', 0x043E));
    repl.insert(std::pair<short, short>(L'p', 0x0440));
    repl.insert(std::pair<short, short>(L'x', 0x0445));
    repl.insert(std::pair<short, short>(L'y', 0x0443));
    repl.insert(std::pair<short, short>(0xF6, 0x04E7));
    repl.insert(std::pair<short, short>(0xFF, 0x04F1));
    repl.insert(std::pair<short, short>(0x04B7, 0x04CC));

    char locale_name[32] = "";
    locale_name[0] = '.';
    _ui64toa_s(1251, locale_name + 1, sizeof(locale_name) - 1, 10);
    locinfo = _create_locale(LC_ALL, locale_name);

    return *hRes;
}

STDMETHODIMP CKhParser::Terminate( long* hRes )
{
    saveResults();
    if (pIXMLHTTPRequest != NULL)
        pIXMLHTTPRequest.Release();
    repl.clear();
    if (locinfo != 0) {
        _free_locale(locinfo);
        locinfo = 0;
    }
    homonyms.clear();
    currHom = -1;
    cache.clear();
    sentences.clear();
    empty.clear();
    safeArraySize = 0;
    request = L"/suddenly/?parse=";

    return S_OK;
}

HRESULT CKhParser::DoParse(BSTR word, long* hRes)
{
    if (word == NULL)
        return S_OK;

    *hRes = S_OK;

    BSTR normWord = NULL;
    this->normWord(word, normWord);
    addWordsToSent(word, normWord);
    addToSentSize(1);

    std::map<std::wstring, int>::iterator eit = empty.find(normWord);
    if (eit != empty.end()) {
        homonyms.clear();
        currHom = -1;
        eit->second++;
    //    addToSentSize(1);
        return S_OK;
    }
    HomMap::iterator it = cache.find(normWord);

    if (it != cache.end()) {
        homonyms = it->second.homVct;
        currHom = 0;
        it->second.count++;
        addToSentSize((int)homonyms.size() - 1);
        return S_OK;
    }
    wchar_t* toPost = new wchar_t [wcslen(request) + wcslen(normWord) + 1];
    memset(toPost, 0,( wcslen(request) + wcslen(normWord) + 1) * sizeof(wchar_t));
    wcscat_s(toPost, wcslen(request) + wcslen(normWord) + 1, request);
    wcscat_s(toPost, wcslen(request) + wcslen(normWord) + 1, normWord);

    *hRes = pIXMLHTTPRequest->open("GET", toPost, false);
    if (*hRes != S_OK) {
        delete [] toPost;
        return *hRes;
    }

    *hRes = pIXMLHTTPRequest->send();
    if (*hRes != S_OK) {
        delete [] toPost;
        return *hRes;
    }

    size_t statuc = pIXMLHTTPRequest->status;
    statuc++;
    if (pIXMLHTTPRequest->status == 200)
    { 
        fillHomonyms(pIXMLHTTPRequest->responseText);
        if (homonyms.size() > 0) {
            cacheItem item;
            item.count = 1;
            item.homVct = homonyms;
            currHom = 0;
            cache.insert(std::pair<std::wstring, cacheItem>(normWord, item));
            addToSentSize((int)homonyms.size() - 1);
        }
        else {
            empty.insert(std::pair<std::wstring, int>(normWord, 1));
//            addToSentSize(1);
        }
    } 
    delete [] toPost;
    return *hRes;
}

HRESULT CKhParser::GetHomonymNum( int* cNumHom )
{
    *cNumHom = (int)homonyms.size();
    return S_OK;
}

HRESULT CKhParser::GetNextHomonym( SAFEARRAY** lpHomonym )
{
    if (currHom == -1)
        return S_OK;
    std::wstring str;
    if (currHom < homonyms.size())
        str = homonyms[currHom].khak;
    else
        str = homonyms[currHom % homonyms.size()].rus;

    long idx, len, start;
    len = (long)str.length();
    SafeArrayLock(*lpHomonym);

    SafeArrayGetUBound(*lpHomonym, 1, &len);
    SafeArrayGetLBound(*lpHomonym, 1, &start);
    short asterisk = L'*';
    for (idx = start; idx < len; idx ++) {
        if (idx == str.length()) {
//        if (str[idx] == 0x0) {
            SafeArrayPutElement(*lpHomonym, &idx, &asterisk);
            break;
        }
        SafeArrayPutElement(*lpHomonym, &idx, &str[ idx ]);
    }
    SafeArrayUnlock(*lpHomonym);

    currHom++;

    return S_OK;
}

HRESULT CKhParser::Normalize( BSTR InputWord, SAFEARRAY** lpHomonym )
{
    return S_OK;
}

HRESULT CKhParser::AddKhakSent(BSTR InputSent, long* hRes)
{
    *hRes = S_OK;
    sent newSent;
    newSent.khak_sent = std::wstring(InputSent);
    newSent.informant = std::wstring(L"");
    size_t pos = newSent.khak_sent.find(0x1f);
    while (pos != std::wstring::npos) {
        newSent.khak_sent.replace(pos, 1, L"");
        pos = newSent.khak_sent.find(0x1f);
    }
    newSent.size = 0;
    sentences.push_back(newSent);
    return S_OK;
}

HRESULT CKhParser::AddKhakSent2(BSTR Name, BSTR InputSent, long* hRes)
{
    *hRes = S_OK;
    sent newSent;
    newSent.khak_sent = std::wstring(InputSent);
    newSent.informant = std::wstring(Name);
    size_t begin = 0, end = newSent.informant.length();
    for (std::wstring::iterator it = newSent.informant.begin(); it != newSent.informant.end(); ++it) {
        if (*it == L' ' || *it == L'\x09')
            begin++;
        else
            break;
    }
    for (std::wstring::iterator it = newSent.informant.end() - 1; it != newSent.informant.begin(); --it) {
        if (*it == L' ' || *it == L'\t')
            end--;
        else
            break;
    }
    newSent.informant = newSent.informant.substr(begin, end);
    if (names.find(newSent.informant) == names.end())
        names.insert(std::pair<std::wstring, int>(newSent.informant, (int)(names.size() + 1)));
    size_t pos = newSent.khak_sent.find(0x1f);
    while (pos != std::wstring::npos) {
        newSent.khak_sent.replace(pos, 1, L"");
        pos = newSent.khak_sent.find(0x1f);
    }
    newSent.size = 0;
    sentences.push_back(newSent);
    return S_OK;
}

HRESULT CKhParser::AddRusSent(BSTR InputSent, long* hRes)
{
    *hRes = S_OK;
    if (sentences.size() > 0) {
        SentVct::iterator it = sentences.end() - 1;
        it->rus_sent = std::wstring(InputSent);
        size_t pos = it->rus_sent.find(0x1f);
        while (pos != std::wstring::npos) {
            it->rus_sent.replace(pos, 1, L"");
            pos = it->rus_sent.find(0x1f);
        }
    }
    return S_OK;
}

HRESULT CKhParser::normWord( const BSTR& InputWord, BSTR& normWord )
{
    size_t len = wcslen(InputWord);
    normWord = new wchar_t [len + 1];
    normWord[0] = 0x0;
    wcscat_s(normWord, len + 1, InputWord);

    _wcslwr_s_l(normWord, len + 1, locinfo);
    for (int i = 0; i < len; i++)
    {
        std::map<short, short>::iterator it = repl.find(normWord[i]);
        if (it != repl.end())
            normWord[i] = it->second;
    }
    for (int i = (int)wcslen(normWord) - 1; i >= 0; i--)
        if (normWord[i] == L' ')
            normWord[i] = 0x0;
    return S_OK;
}

HRESULT CKhParser::fillHomonyms(BSTR response)
{
    wchar_t* str(response);
    int i = 0;
    const wchar_t* foundStem = L"FOUND STEM:";
    homonyms.clear();
    hom homonym;
    wchar_t* tmp = wcsstr(str, foundStem);
    while (tmp != 0) {
        tmp = tmp + 12;
        // grammatics
        wchar_t* form = getDetails(tmp, L' ');
        // affixes
        wchar_t* affixes = getDetails(wcschr(tmp, L' '), 0xA);

        tmp = wcschr(tmp, 0xA);
        if (tmp == 0)
            continue;
        //part of speech
        wchar_t* pos = getSubstr(tmp, L' ');
        tmp = wcschr(tmp, L' ');
        if (tmp == 0)
            continue;
        // headword
        wchar_t* headword = getSubstr(tmp, L' ');
        tmp = wcschr(tmp, 0x201B);
        if (tmp == 0)
            continue;
        tmp = tmp + 1;
        // meaning
        wchar_t* meaning = getSubstr(tmp, 0x2019);
        tmp = wcschr(tmp, 0x2019) + 1;
        wchar_t* stem = getSubstr(tmp, 0xA);
        if (stem[0] != 0x0) {
            for (i = (int)wcslen(stem) - 1; i >= 0; i--)
                if (stem[i] == L' ')
                    stem[i] = 0x0;
        }
        if (stem[0] != 0x0)
            homonym.khak = std::wstring(stem);
        else
            homonym.khak = std::wstring(headword);
        homonym.khak.append(affixes);
        homonym.rus = std::wstring(meaning).append(form);
        homonym.lemma = std::wstring(headword);
        homonym.pos = std::wstring(pos);

        homonyms.push_back(homonym);
        tmp = tmp = wcsstr(tmp, foundStem);
     }
    return S_OK;
}

wchar_t* CKhParser::getDetails(const wchar_t* str, wchar_t endCh)
{
    wchar_t* aff = 0;
    int len = 0;
    int i = 0, j = 0;
    char dash = 0;
   // str = Trim$(str)
    for (i = 0; str[i] == ' '; i++);
    int start = i;
    for (i; str[i] != endCh; i++) {
        if (str[i] != L'-') {
            if (dash == 1)
                len ++;
            len++;
            dash = 0;
        }
        else
            dash = 1;
    }
    aff = new wchar_t [len + 1];
    i = start;
    dash = 0;
    j = 0;
    while (str[i] != endCh) {
        if (str[i] != L'-') {
            if (dash == 1)
                aff[j++] = L'-';
            aff[j++] = str[i++];
            dash = 0;
        }
        else {
            dash = 1;
            i++;
        }
    }
    aff[j] = 0x0;
    return aff;
}

wchar_t* CKhParser::getSubstr(const wchar_t* str, wchar_t endCh)
{
    wchar_t* headword = 0;
    int i, start = 0;
    for (i = 0; str[i] == ' '; i++);
    start = i;
    size_t len = wcschr(str + start, endCh) - (str + start);
    headword = new wchar_t [len + 1];
    memcpy(headword, str + start, len*sizeof(wchar_t));
    headword[len] = 0x0;
    return headword;
}

HRESULT CKhParser::saveResults(void)
{
//    char locale_name[32] = "";
//    locale_name[0] = '.';
//    _ui64toa_s(1251, locale_name + 1, sizeof(locale_name) - 1, 10);
    std::locale loc = std::locale(std::locale("C"), new std::codecvt_utf8<wchar_t, 0x10ffff, std::generate_header>());
    // print words parsed
    if (dict.length() != 0 && cache.size() != 0) {
        std::wofstream outDict(dict, std::wofstream::binary);
        if (outDict.is_open()) {
//            std::locale loc;
//            loc = std::locale("ru_RU.utf8");
            outDict.imbue(loc);

//            outDict.imbue(loc);//locale_name));
            HomMap::iterator it = cache.begin();
            for (; it != cache.end(); ++it) {
                outDict.write(it->first.c_str(), it->first.length());// * sizeof(wchar_t));
                outDict.write(L":\t", wcslen(L":\t"));// * sizeof(wchar_t));
                wchar_t count[10] = L"";
                _itow_s(it->second.count, count, 10, 10);
//                wsprintf((LPSTR)count, (LPCSTR)L"%d", it->second.count);
                outDict.write(count, wcslen(count));// * sizeof(wchar_t));// << L": \t" << it->second.count << L"\t";
                outDict.write(L":\t", wcslen(L":\t"));// * sizeof(wchar_t));
                HomVct::iterator vt = it->second.homVct.begin();
                for (; vt != it->second.homVct.end(); ++vt) {
                    outDict.write(vt->khak.c_str(), vt->khak.length());// * sizeof(wchar_t));
                    outDict.write(L" : ", wcslen(L" : "));// * sizeof(wchar_t));
                    outDict.write(vt->rus.c_str(), vt->rus.length());// * sizeof(wchar_t));
                    outDict.write(L"; ", wcslen(L"; "));// * sizeof(wchar_t));
    //                outDict << vt->khak << L" : " << vt->rus << L" ; ";
                }
                outDict.write(L"\n", wcslen(L"\n"));// * sizeof(wchar_t));
    //            outDict << L"\n";
            }
        }
        outDict.close();
    }
    // print words not found
    if (notfound.length() != 0 && empty.size() != 0) {
        std::wofstream notFound(notfound, std::wofstream::binary);
        if (notFound.is_open()) {
            notFound.imbue(loc);
            std::map<std::wstring, int>::iterator et = empty.begin();
            for (; et != empty.end(); ++et) {
                notFound.write(et->first.c_str(), et->first.length());
                notFound.write(L" : ", wcslen(L" : "));// * sizeof(wchar_t));
                wchar_t count[10] = L"";
                _itow_s(et->second, count, 10, 10);
//                wsprintf((LPSTR)count, (LPCSTR)L"%d", et->second);
                notFound.write(count, wcslen(count));// * sizeof(wchar_t));// << L": \t" << it->second.count << L"\t";
                notFound.write(L"\n", wcslen(L"\n"));// * sizeof(wchar_t));
//                notFound << et->first << L" : " << et->second << L"\n";
            }
            notFound.close();
        }
    }
    return S_OK;
}

void CKhParser::addToSentSize(int value)
{
    if (sentences.size() > 0) {
        SentVct::iterator it = sentences.end() - 1;
        it->size += value;
    }
}

void CKhParser::addWordsToSent(BSTR word, BSTR normWord)
{
    if (sentences.size() > 0) {
        SentVct::iterator it = sentences.end() - 1;
        std::map<std::wstring, std::wstring>::iterator kt = it->keys.find(word);
        if (kt == it->keys.end()) {
            it->keys.insert(std::pair<std::wstring, std::wstring>(word, normWord));
        }
        it->words.push_back(word);
    }
}

