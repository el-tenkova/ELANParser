// CKhParser.cpp : Implementation of CKhParser
#ifdef _WINDOWS
    #include "stdafx.h"
#endif
#include <string>
#include <locale.h>
#include <fstream>
#include <codecvt>
#include <sstream>
#include <algorithm>
#include <iterator>
#include <functional>
#include <iostream>

#include "CKhParser.h"

#ifdef _WINDOWS
#define RUS_LOCALE "Russian"
#endif

std::wstring CKhParser::Kh_Sent = L"KH_Sent";
std::wstring CKhParser::Transcr = L"Transcr";
std::wstring CKhParser::Kh_Words = L"Kh_Words";
std::wstring CKhParser::Kh_Homonyms = L"Kh_Homonyms";
std::wstring CKhParser::Kh_Lemma = L"Kh_Lemma";
std::wstring CKhParser::Kh_PartOfSpeech = L"Kh_PartOfSpeech";
std::wstring CKhParser::Kh_Morphems = L"Kh_Morphems";
std::wstring CKhParser::Rus_Morphems = L"Rus_Morphems";
std::wstring CKhParser::Eng_Morphems = L"Eng_Morphems";
std::wstring CKhParser::Rus_Homonyms = L"Rus_Homonyms";
std::wstring CKhParser::Eng_Homonyms = L"Eng_Homonyms";
std::wstring CKhParser::Rus_Sent = L"Rus_Sent";

CKhParser::CKhParser()
:pIXMLHTTPRequest(NULL)
,request(L"/suddenly/?parse=")
#ifdef _WINDOWS
,locinfo(0)
#endif
,homonyms(0)
,currHom(-1)
,dict("")
,notfound("")
,sentences(0)
,russian(std::locale(RUS_LOCALE), new std::codecvt_utf8<wchar_t, 0x10ffff, std::consume_header>())
{}
#ifdef _WINDOWS
long CKhParser::Init(const std::wstring& www, const std::string& dictPath, const std::string& notfoundPath)
#else
long CKhParser::Init(const std::string& www, const std::string& dictPath, const std::string& notfoundPath)
#endif
{
    Terminate();
    long hRes = 0;
    dict = std::string(dictPath);
    notfound = std::string(notfoundPath);
    //"http://khakas.altaica.ru"
#ifdef _WINDOWS
    request = www + request;
    _locale_t locale = _create_locale(LC_ALL, "ru-RU");
    char tmp[512] = "";
    size_t len = 0;
    _wcstombs_s_l(&len, tmp, sizeof(tmp), www.c_str(), _TRUNCATE, locale);
    std::string _www(tmp);
    pIXMLHTTPRequest = new CKhHttpWrapper(&hRes, _www.c_str(), "");
    _free_locale(locale);
#else
    size_t pos = www.find("://");
    std::string protocol(pos != std::string::npos ? www.substr(0, www.find("://")) : "http");
    std::cout << "protocol = " << protocol << std::endl;
    std::wstring www1(www.begin(), www.end());
    request = www1 + request;
    std::string www2(pos != std::string::npos ? www.substr(protocol.length() + 3) : www);
    pIXMLHTTPRequest = new CKhHttpWrapper(&hRes, www2.c_str(), protocol.c_str());
    std::cout << "www = " << www << std::endl;
    std::cout << "www2 = " << www2 << std::endl;
    std::wcout << L"request = " << request << std::endl;
#endif
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

#ifdef _WINDOWS
    char locale_name[32] = "";
    locale_name[0] = '.';
    _ui64toa_s(1251, locale_name + 1, sizeof(locale_name) - 1, 10);
    locinfo = _create_locale(LC_ALL, locale_name);
#endif
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Kh_Sent, L"Transcription-txt-kjh"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Transcr, L"Phonetic-txt-kjh"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Kh_Words, L"Words-txt-kjh"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Kh_Homonyms, L"Morph-txt-kjh"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Kh_Lemma, L"Lemma-txt-kjh"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Kh_PartOfSpeech, L"POS-txt-en"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Kh_Morphems, L"mb"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Rus_Morphems, L"gr"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Eng_Morphems, L"ge"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Rus_Homonyms, L"Gloss-txt-rus"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Eng_Homonyms, L"Gloss-txt-en"));
    lvlNames.insert(std::pair<std::wstring, std::wstring>(Rus_Sent, L"Translation-gls-rus"));
    return hRes;
}

long CKhParser::Terminate(void)
{
#ifdef _WINDOWS
    saveResults();
#endif
    delete pIXMLHTTPRequest;
    repl.clear();
#ifdef _WINDOWS
    if (locinfo != 0) {
        _free_locale(locinfo);
        locinfo = 0;
    }
#endif
    homonyms.clear();
    currHom = -1;
    cache.clear();
    sentences.clear();
    empty.clear();
    request = L"/suddenly/?parse=";

    return 0;
}

long CKhParser::DoParse(const std::wstring& word)
{
    if (word.empty())
        return 0;

    long hRes = 0;

    std::wstring normWord;
    this->normWord(word, normWord);
    addWordsToSent(word, normWord);
    addToSentSize(1);

    std::map<std::wstring, int>::iterator eit = empty.find(normWord);
    if (eit != empty.end()) {
        homonyms.clear();
        currHom = -1;
        eit->second++;
        //    addToSentSize(1);
        return 0;
    }
    HomMap::iterator it = cache.find(normWord);

    if (it != cache.end()) {
        homonyms = it->second.homVct;
        currHom = 0;
        it->second.count++;
        size_t hSize = 0;
        for (auto it = homonyms.begin(); it != homonyms.end(); ++it)
            hSize += it->morphems.size();
        addToSentSize((int)hSize - 1); //homonyms.size() - 1);
        return 0;
    }

    std::wstring toPost;
    toPost.append(request);
    toPost.append(normWord);
    hRes = pIXMLHTTPRequest->Open(toPost.c_str());
    if (hRes != 0)
        return hRes;

    hRes = pIXMLHTTPRequest->Send();
    if (hRes != 0)
        return hRes;

    size_t status= pIXMLHTTPRequest->Status();
    if (status == 200)
    { 
        fillHomonyms(pIXMLHTTPRequest->ResponseText());
        if (homonyms.size() > 0) {
            cacheItem item;
            item.count = 1;
            item.homVct = homonyms;
            currHom = 0;
            cache.insert(std::pair<std::wstring, cacheItem>(normWord, item));
            size_t hSize = 0;
            for (auto it = homonyms.begin(); it != homonyms.end(); ++it)
                hSize += it->morphems.size();
            addToSentSize((int)hSize - 1); //homonyms.size() - 1);
        }
        else {
            empty.insert(std::pair<std::wstring, int>(normWord, 1));
//            addToSentSize(1);
        }
    } 
    return hRes;
}

int CKhParser::GetHomonymNum(void)
{
    return ((int)homonyms.size());
}

long CKhParser::GetNextHomonym(std::wstring& hom)
{
    if (currHom == -1)
        return 0;
    if (currHom < homonyms.size())
        hom = homonyms[currHom].khak;
    else
        hom = homonyms[currHom % homonyms.size()].rus;

    currHom++;
    return 0;
}

long CKhParser::AddKhakSent(const std::wstring& InputSent)
{
    sent newSent;
    newSent.khak_sent = std::wstring(InputSent);
    newSent.informant = std::wstring(L"");
    newSent.begin = 0;
    newSent.end = 0;
    size_t pos = newSent.khak_sent.find(0x1f);
    while (pos != std::wstring::npos) {
        newSent.khak_sent.replace(pos, 1, L"");
        pos = newSent.khak_sent.find(0x1f);
    }
    newSent.size = 0;
    sentences.push_back(newSent);
    return 0;
}

long CKhParser::AddKhakSent2(const std::wstring& Name, const std::wstring& InputSent)
{
    sent newSent;
    newSent.khak_sent = std::wstring(InputSent);
    newSent.informant = std::wstring(Name);
    newSent.begin = 0;
    newSent.end = 0;
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
    return 0;
}

long CKhParser::AddKhakSent3(const std::wstring& Name, const std::wstring& Time, const std::wstring& InputSent)
{
    if (InputSent.length() > 0)
        lvlExist[Kh_Sent] = 1;
    sent newSent;
    newSent.khak_sent = std::wstring(InputSent);
    newSent.informant = Name.length() > 0 ? std::wstring(Name) : std::wstring(L"I");
    newSent.begin = 0;
    newSent.end = 0;
    std::wstring beg(Time);
    int seconds = 0, minutes = 0;
    if (beg.length() > 0)
    {
        minutes = std::stoi(beg);
        size_t pt = beg.find(L'.');
        if (pt != std::wstring::npos && pt < beg.length() - 1)
            seconds = std::stoi(beg.substr(pt + 1));
    }
    newSent.begin = (minutes * 60 + seconds) * 1000;
    if (sentences.size() > 0 && newSent.begin == 0)
        newSent.begin = sentences.size() * 5000;
    size_t begin = 0, end = newSent.informant.length();
    for (std::wstring::iterator it = newSent.informant.begin(); it != newSent.informant.end(); ++it) {
        if (*it == L' ' || *it == L'\x09')
            begin++;
        else
            break;
    }
    if (newSent.informant.length() > 0)
    {
        for (std::wstring::iterator it = newSent.informant.end() - 1; it != newSent.informant.begin(); --it) {
            if (*it == L' ' || *it == L'\t')
                end--;
            else
                break;
        }
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
    return 0;
}

long CKhParser::AddRusSent(const std::wstring& InputSent)
{
    if (InputSent.length() > 0)
        lvlExist[Rus_Sent] = 1;
    if (sentences.size() > 0) {
        SentVct::iterator it = sentences.end() - 1;
        it->rus_sent = std::wstring(InputSent);
        size_t pos = it->rus_sent.find(0x1f);
        while (pos != std::wstring::npos) {
            it->rus_sent.replace(pos, 1, L"");
            pos = it->rus_sent.find(0x1f);
        }
    }
    return 0;
}

long CKhParser::AddTranscription(const std::wstring& InputSent)
{
    if (InputSent.length() > 0)
        lvlExist[Transcr] = 1;
    if (sentences.size() > 0) {
        SentVct::iterator it = sentences.end() - 1;
        it->transcr = std::wstring(InputSent);
        size_t pos = it->transcr.find(0x1f);
        while (pos != std::wstring::npos) {
            it->transcr.replace(pos, 1, L"");
            pos = it->transcr.find(0x1f);
        }
    }
    return 0;
}

long CKhParser::normWord(const std::wstring& InputWord, std::wstring& normWord )
{
    size_t len = InputWord.length();
//    normWord = new wchar_t [len + 1];
//    normWord[0] = 0x0;
//    wcscat_s(normWord, len + 1, InputWord);
    normWord.clear();
    normWord.append(InputWord);
    std::transform(normWord.begin(), normWord.end(), normWord.begin(),
        std::bind2nd(std::ptr_fun(&std::tolower<wchar_t>), russian));
//    _wcslwr_s_l(normWord, len + 1, locinfo);
    for (int i = 0; i < len; i++)
    {
        std::map<short, short>::iterator it = repl.find(normWord[i]);
        if (it != repl.end())
            normWord[i] = it->second;
    }
    for (int i = (int)normWord.length() - 1; i >= 0; i--)
        if (normWord[i] == L' ')
            normWord[i] = 0x0;
    return 0;
}

long CKhParser::fillHomonyms(const std::wstring& response)
{
    std::wstring tmpResp(response);
    wchar_t* str = new wchar_t [response.length() + 1];
    str[0] = 0;
    wcscat(str, response.c_str());
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
        wchar_t* pos = getSubstr(tmp + 1, L' ');
        tmp = wcschr(tmp, L' ');
        if (tmp == 0)
            continue;
        // headword
        wchar_t* headword = getSubstr(tmp, L' ');
        wchar_t* poss = wcsstr(tmp, L"<+poss>");
        tmp = wcschr(tmp, 0x201B);
        if (tmp == 0)
            continue;
        if (poss != 0 && poss > tmp)
            poss = 0;
        tmp = tmp + 1;
        // meaning
        wchar_t* meaning = getSubstr(tmp, 0x2019);
        tmp = wcschr(tmp, 0x2019) + 1;
        wchar_t* stem = getSubstr(tmp, 0xA);
        if (stem[0] >= L'0' && stem[0] <= L'9')
            stem[0] = 0x0;
        if (stem[0] != 0x0) {
            for (i = (int)wcslen(stem) - 1; i >= 0; i--)
                if (stem[i] == L' ')
                    stem[i] = 0x0;
        }
        if (stem[0] != 0x0 && stem[0] != L'-')
            homonym.khak = std::wstring(stem);
        else if (stem[0] == L'-')
            homonym.khak = std::wstring(L"stem");
        else
            homonym.khak = std::wstring(headword);
        homonym.morphems.push_back(homonym.khak);
        getMorphems(affixes, homonym.morphems, L'-');
        homonym.khak.append(affixes);

        homonym.rus = std::wstring(meaning);
        if (poss != 0)
            homonym.rus.append(L".3pos");
        form = removeSymbols(form);
        homonym.rus.append(form);

        homonym.r_morphems.push_back(meaning);
        getMorphems(form, homonym.r_morphems, L'-');
        homonym.lemma = std::wstring(headword);
        pos = removeSymbols(pos);
        homonym.pos = std::wstring(pos);

        homonyms.push_back(homonym);
        homonym.morphems.clear();
        homonym.r_morphems.clear();
        tmp = wcsstr(tmp, foundStem);
     }
    delete [] str;
    return 0;
}

wchar_t* CKhParser::removeSymbols(wchar_t* input)
{
    size_t len = wcslen(input);
    const wchar_t* symbols[][2] = 
    {
        {L"\x1d62", L""},
        {L"\x1d63", L""},
        {L"\x2090", L""},
        {L"\x2080", L""},
        //
        {L"Ass\x2082", L"Ass2" },
        { L"i\x2081", L"i1" },
        {L"Gen\x2081", L"Gen1" },
        {L"\x2081", L""},
        {L"\x2082", L""},
        {L"\x2083", L""},
    };
    for (size_t i = 0; i < sizeof(symbols) / sizeof(symbols[0]); i++) {
        wchar_t* pos = wcsstr(input, symbols[i][0]);
        while (pos != 0) {
            size_t delta = wcslen(symbols[i][0]) - wcslen(symbols[i][1]);
            memcpy(pos, symbols[i][1], wcslen(symbols[i][1]) * sizeof(wchar_t));
            memcpy(pos + wcslen(symbols[i][1]), pos + wcslen(symbols[i][0]), wcslen(pos + wcslen(symbols[i][1])) * sizeof(wchar_t));
           /// input[wcslen(input) - 1 - delta] = 0x0;
            pos = wcsstr(input, symbols[i][0]);
        }
    }
    return input;
}

wchar_t* CKhParser::getDetails(const std::wstring& str, wchar_t endCh)
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
    if (wcschr(str + start, endCh) == 0)
    {
        headword = new wchar_t[2];
        headword[0] = L'-';
        headword[1] = 0x0;
        return headword;
    }
    size_t len = wcschr(str + start, endCh) - (str + start);
    headword = new wchar_t [len + 1];
    memcpy(headword, str + start, len*sizeof(wchar_t));
    headword[len] = 0x0;
    return headword;
}

long CKhParser::saveResults(void)
{
//    char locale_name[32] = "";
//    locale_name[0] = '.';
//    _ui64toa_s(1251, locale_name + 1, sizeof(locale_name) - 1, 10);
//    std::locale loc = std::locale(std::locale("C"), new std::codecvt_utf8<wchar_t, 0x10ffff, std::generate_header>());
    // print words parsed
    if (dict.length() != 0 && cache.size() != 0) {
        std::wofstream outDict(dict, std::wofstream::binary);
        if (outDict.is_open()) {
//            std::locale loc;
//            loc = std::locale("ru_RU.utf8");
            outDict.imbue(russian);

//            outDict.imbue(loc);//locale_name));
            HomMap::iterator it = cache.begin();
            for (; it != cache.end(); ++it) {
                outDict.write(it->first.c_str(), it->first.length());// * sizeof(wchar_t));
                outDict.write(L":\t", wcslen(L":\t"));// * sizeof(wchar_t));
                //int n = it->second.count;
                std::wstring nstr = std::to_wstring(it->second.count);
                outDict.write(nstr.c_str(), nstr.length());
                //wchar_t count[10] = L"";
                //_itow_s(it->second.count, count, 10, 10);
//                wsprintf((LPSTR)count, (LPCSTR)L"%d", it->second.count);
                //outDict.write(count, wcslen(count));// * sizeof(wchar_t));// << L": \t" << it->second.count << L"\t";
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
            notFound.imbue(russian);
            std::map<std::wstring, int>::iterator et = empty.begin();
            for (; et != empty.end(); ++et) {
                notFound.write(et->first.c_str(), et->first.length());
                notFound.write(L" : ", wcslen(L" : "));// * sizeof(wchar_t));
                std::wstring nstr = std::to_wstring(et->second);
                notFound.write(nstr.c_str(), nstr.length());
//                notFound.write(std::endl);
                //wchar_t count[10] = L"";
                //_itow_s(et->second, count, 10, 10);
//                wsprintf((LPSTR)count, (LPCSTR)L"%d", et->second);
                //notFound.write(count, wcslen(count));// * sizeof(wchar_t));// << L": \t" << it->second.count << L"\t";
                notFound.write(L"\n", wcslen(L"\n"));// * sizeof(wchar_t));
//                notFound << et->first << L" : " << et->second << L"\n";
            }
            notFound.close();
        }
    }
    return 0;
}

void CKhParser::addToSentSize(int value)
{
    if (sentences.size() > 0) {
        SentVct::iterator it = sentences.end() - 1;
        it->size += value;
    }
}

void CKhParser::addWordsToSent(const std::wstring& word, const std::wstring& normWord)
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

void CKhParser::getMorphems(const std::wstring& str, std::vector<std::wstring>& morphems, const wchar_t& delim)
{
    size_t pos = str.find(delim);
    if (pos == std::wstring::npos)
        return;
    size_t prev = 0;
    while (pos != std::wstring::npos) {
        if (pos == 0) {
            prev = pos + 1;
            pos = str.find(delim, pos + 1);
            continue;
        }
        morphems.push_back(str.substr(prev, pos - prev));
        prev = pos + 1;
        if (pos + 1 > str.length())
            break;
        pos = str.find(delim, pos + 1);
    }
    morphems.push_back(str.substr(prev));
}

size_t CKhParser::getMorphemsCount(HomMap::iterator& ct)
{
    size_t s = 0;
    for (auto hit = ct->second.homVct.begin(); hit != ct->second.homVct.end(); ++hit) {
        s += hit->morphems.size();
    }
    return s;
}
