// KhParser.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#include <string>
#include <vector>
#include <map>

using namespace std;

#import "msxml6.dll"
using namespace MSXML2;

void ParseResponse(_bstr_t respStr, map<wstring, vector<map<wstring, wstring>>> words)
{
    long i = 0;
    long j = 0;
    long k = 0;
    long idx = 1;
    wchar_t* tmp = 0;
//    Dim str As String
 //   Dim result As New Collection
//    Dim idx As Long
//    idx = 1
    
    wchar_t* str = (wchar_t*)respStr;
//    str = objHTTP.responsetext
    
    tmp = wcsstr(str, L"FOUND STEM:");
    vector<map<wstring, wstring>> homonyms;
    while (tmp != 0)
    {
 //       Set homonym = New Collection
        map<wstring, wstring> hom;
        str = tmp + 12;
//        str = Mid(str, i + 12)
        j = wcschr(str, L' ') - str;
        i = wcschr(str, L'\x10') - str;
//        'Debug.Print str
        hom.insert(std::pair<wstring, wstring>(L"form", wstring(str, 1, j)));
//        homonym.Add Mid(str, 1, j), "form"
        hom.insert(std::pair<wstring, wstring>(L"affixes", wstring(str, j + 1, i - (j + 1))));
//        homonym.Add Mid(str, j + 1, i - (j + 1)), "affixes"
        str = str + i + 1;
        hom.insert(std::pair<wstring, wstring>(L"p_o_s", wstring(str, 1, 1)));
//        homonym.Add Mid(str, 1, 1), "p_o_s"
        i = wcschr(str, L'\x10') - str;

        wchar_t* dict = str + 2;//, i - 2);
        tmp = wcsstr(dict, L" \u201B");
        j = tmp - dict;
        hom.insert(std::pair<wstring, wstring>(L"headword", wstring(dict, 1, j)));
//        homonym.Add Mid(dict, 1, j), "headword"
        k = wcschr(dict, L'\u2019') - str;
        hom.insert(std::pair<wstring, wstring>(L"meaning", wstring(dict,j + 1, k + 1 - (j + 1))));
//        homonym.Add Mid(dict, j + 1, k + 1 - (j + 1)), "meaning"

//        homonym.Add "0", "has-stem"
        hom.insert(std::pair<wstring, wstring>(L"has-stem", L"0"));
        if (wcslen(dict) > k + 1) {
            hom.insert(std::pair<wstring, wstring>(L"hasstem", wstring(dict, k + 2)));
//            homonym.Add Mid(dict, k + 2), "stem"
            hom[L"has-stem"] = wstring(L"1");
//            homonym.Remove "has-stem"
//            homonym.Add "1", "has-stem"
        }

        homonyms.push_back(hom);
//        idx = idx + 1;

        tmp = wcsstr(tmp, L"FOUND STEM:");
    }
    
}

string Utf8_to_cp1251(const char *str)
{
	string res;
	int result_u, result_c;

	result_u = MultiByteToWideChar(CP_UTF8,
		0,
		str,
		-1,
		0,
		0);

	if (!result_u)
		return 0;

	wchar_t *ures = new wchar_t[result_u];

	if(!MultiByteToWideChar(CP_UTF8,
		0,
		(LPCSTR)str,
		-1,
		ures,
		result_u))
	{
		delete[] ures;
		return 0;
	}
    
	result_c = WideCharToMultiByte(
		CP_UTF8,
		0,
		ures,
		-1,
		0,
		0,
		0, 0);

	if(!result_c)
	{
		delete [] ures;
		return 0;
	}

	char *cres = new char[result_c];

	if(!WideCharToMultiByte(
		1251,
		0,
		ures,
		-1,
		cres,
		result_c,
		0, 0))
	{
		delete[] cres;
		return 0;
	}
	delete[] ures;
	res.append(cres);
	delete[] cres;
	return res;
}
int _tmain(int argc, _TCHAR* argv[])
{
    BSTR bstrString = NULL;
    CoInitialize(NULL);
    try {
        IXMLHTTPRequestPtr pIXMLHTTPRequest = NULL;
        HRESULT hr;
        
        hr=pIXMLHTTPRequest.CreateInstance("Msxml2.ServerXMLHTTP");
        SUCCEEDED(hr) ? 0 : throw hr;

        _bstr_t req = "http://khakas.altaica.ru/cgi-bin/suddenly.fs?Getparam=getval&parse=";
        //wchar_t in[1024] = L"";
        char in[1024] = "";
        for (int i = 0; i < 4; i++)
            fgetc(stdin);
  //      vector<wstring> khak;
        vector<string> khak;
        while (!feof(stdin)) {
           // memset(in, 0, sizeof(in));
            // get article
            //fgetws((wchar_t*)in, sizeof(in), stdin);
            fgets((char*)in, sizeof(in), stdin);
//            if (wcslen(in) == 0)
            if (strlen(in) == 0)
                continue;
//            wchar_t dashes[] = L"-----";
            char dashes[] = "-----";
//            if (wcsncmp(in, dashes, wcslen(dashes)) == 0) {
            if (strncmp(in, dashes, strlen(dashes)) == 0) {
                if (khak.size() > 0) {
//                    vector<wstring>::iterator it = khak.begin() + 1;
                    vector<string>::iterator it = khak.begin() + 1;
                    for (; it != khak.end() - 1; ++it) {
                        if ((*it)[0] == 'K' || (*it)[0] == 'R')
                            continue;
                        _bstr_t toPost(req);
                        toPost = toPost + ((*it).c_str());
                        hr=pIXMLHTTPRequest->open("POST", toPost, false);
                        SUCCEEDED(hr) ? 0 : throw hr;
    
                      //  hr=pIXMLHTTPRequest->setRequestHeader("Content-Type", "text/plain;charset=UTF-8");
                      //  SUCCEEDED(hr) ? 0 : throw hr;

                        hr=pIXMLHTTPRequest->send();
                        SUCCEEDED(hr) ? 0 : throw hr;

                        if (pIXMLHTTPRequest->status)
                        {
                            bstrString=pIXMLHTTPRequest->responseText;

                            if(bstrString)
                            {
                                ::SysFreeString(bstrString);
                                bstrString = NULL;
                            }
                        }
                    }
                    khak.clear();
                }
                continue;
            }
//            if (wcschr(in, L'|') != 0)
            if (strchr(in, '|') != 0)
//                khak.push_back(string(Utf8_to_cp1251((const wchar_t*)(wcschr(in, L'|') + 1))));
                khak.push_back(string(Utf8_to_cp1251((const char*)(strchr(in, '|') + 1))));
            else
                khak.push_back(in);
        }
    }
    catch (...) {
        MessageBox(NULL, _T("Exception occurred"), _T("Error"), MB_OK);
        if(bstrString)
            ::SysFreeString(bstrString);
    }
    CoUninitialize();

    return 0;
}

