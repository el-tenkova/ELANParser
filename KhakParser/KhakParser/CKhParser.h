
#ifndef __KHPARSER_H_
#define __KHPARSER_H_

#include <atlcom.h>
#include <locale.h>

#import "msxml6.dll"
using namespace MSXML2;

#include <string>
#include <vector>
#include <map>

#include "resource.h"       // main symbols

struct hom
{
    std::wstring khak;
    std::wstring rus;
};

typedef std::vector<hom> HomVct;

struct cacheItem
{
    int count;
    HomVct homVct;
};

typedef std::map<std::wstring, cacheItem> HomMap;

struct sent
{
    int size; // суммарное количество омонимов в предложении (если слово не парсировано +1)
    std::wstring khak_sent; // хакасское предложение
    std::wstring rus_sent; // русское предложение
    std::vector<std::wstring> words; // последовательность слов в предложении
    std::map<std::wstring, std::wstring> keys; // исходное слово : нормализованное слово (ключ в словаре cache)
};

typedef std::vector<sent> SentVct;

class ATL_NO_VTABLE CKhParser : 
    public ATL::CComObjectRootEx<ATL::CComSingleThreadModel>,
    public ATL::CComCoClass<CKhParser, &CLSID_KhParser>,
    public ATL::IDispatchImpl<IKhParser, &IID_IKhParser, &LIBID_KhakParserLib>
{
public:
    CKhParser():
      pIXMLHTTPRequest(NULL),
      request(L"/cgi-bin/suddenly.fs?Getparam=getval&parse="),
      locinfo(0),
      safeArraySize(0),
      homonyms(0),
      currHom(-1),
      dict(L""),
      notfound(L""),
      sentences(0)
/*      tmpResp(L"Parsing: тустарда\n\
FOUND STEM: ----------Pl---Loc---- ----------тар---да----\n\
n тус ‛соль’ \n\
FOUND STEM: ----------Pl---Loc---- ----------тар---да----\n\
n тус ‛время’ \n\
FOUND STEM: ----------Pl---Loc---- ----------тар---да----\n\
n тус ‛напротив’ \n\
3  wordforms generated in 3571 μs.") */
    {
    }

DECLARE_REGISTRY_RESOURCEID(IDR_KHAKPARSER)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CKhParser)
    COM_INTERFACE_ENTRY(IKhParser)
    COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

    STDMETHOD(Init)(int cSafeArr, BSTR www, BSTR dict, BSTR notfound, /*[out, retval]*/ long* hRes);
    STDMETHOD(Terminate)(/*[out, retval]*/ long* hRes);
    STDMETHOD(DoParse)(BSTR InputWord, /*[out, retval] */long* hRes);
    STDMETHOD(GetHomonymNum)(/*[out, retval]*/ int* cNumHom);
    STDMETHOD(GetNextHomonym)(/*[in, out] */SAFEARRAY** lpHomonym);
    STDMETHOD(Normalize)(BSTR InputWord, /*[in, out] */SAFEARRAY** lpNormalized);
    STDMETHOD(AddKhakSent)(BSTR InputSent, /*[out, retval]*/ long* hRes);
    STDMETHOD(AddRusSent)(BSTR InputSent, /*[out, retval]*/ long* hRes);
    STDMETHOD(SaveToELAN)(BSTR ElanPath, /*[out, retval]*/ long *hRes);

protected:
    IXMLHTTPRequestPtr pIXMLHTTPRequest;
    _bstr_t request;
    std::vector<hom> homonyms;
    int currHom;
    HomMap cache;
    SentVct sentences;
    std::map<std::wstring, int> empty;
    std::map<short, short> repl;
    _locale_t locinfo;
    int safeArraySize;

    std::wstring dict;
    std::wstring notfound;

    HRESULT normWord(const BSTR& inputWord, BSTR& normWord);
    HRESULT fillHomonyms(BSTR response);
    wchar_t* getDetails(const wchar_t* str, wchar_t endCh);
    wchar_t* getSubstr(const wchar_t* str, wchar_t endCh);
    HRESULT saveResults(void);
    void addToSentSize(int value);
    void addWordsToSent(BSTR word, BSTR normWord);

    // ELAN part
    void writeHeader(std::wofstream& ef);
    void writeTail(std::wofstream& ef);
    void writeTimeSlot(std::wofstream& ef, const _ULonglong& idx, const _ULonglong& begin);
    void writeTierHeader(std::wofstream& ef, const std::wstring& name, const std::wstring& parent);
    void writeTierTail(std::wofstream& ef);
    void writeAnno(std::wofstream& ef, const std::wstring& sent, const _ULonglong& idx, const _ULonglong& time1, const _ULonglong& time2);
    void writeRefAnno(std::wofstream& ef, const std::wstring& sent, const _ULonglong& idx, const _ULonglong& refid);

    void writeTimes(std::wofstream& ef);
    _ULonglong writeKhakSent(std::wofstream& ef);
    _ULonglong writeRusSent(std::wofstream& ef, _ULonglong id);
    _ULonglong writeWords(std::wofstream& ef, _ULonglong id);
    _ULonglong writeKhakHoms(std::wofstream& ef, _ULonglong id);
    _ULonglong writeRusHoms(std::wofstream& ef, _ULonglong id, _ULonglong refid);

};

#endif //__KHPARSER_H_
