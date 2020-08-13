#ifndef __OROSSPARSER_H_
#define __OROSSPARSER_H_

#include <atlcom.h>
#include <map>

#include "resource.h"       // main symbols

#include "CKhParser.h"

class ATL_NO_VTABLE CKhParserCOM :
    public ATL::CComObjectRootEx<ATL::CComSingleThreadModel>,
    public ATL::CComCoClass<CKhParserCOM, &CLSID_KhParser>,
    public ATL::IDispatchImpl<IKhParser, &IID_IKhParser, &LIBID_KhakParserLib>
{
public:
    CKhParserCOM() : safeArraySize(0)
    {} ;

    DECLARE_REGISTRY_RESOURCEID(IDR_KHAKPARSER)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CKhParserCOM)
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
    STDMETHOD(AddTranscription)(BSTR InputSent, /*[out, retval]*/ long* hRes);
    STDMETHOD(AddKhakSent2)(BSTR Name, BSTR InputSent, /*[out, retval]*/ long* hRes);
    STDMETHOD(AddKhakSent3)(BSTR Name, BSTR Time, BSTR InputSent, /*[out, retval]*/ long* hRes);
    STDMETHOD(SaveToELAN)(BSTR ElanPath, /*[out, retval]*/ long *hRes);
    STDMETHOD(SaveToELANFlex)(BSTR ElanPath, /*[out, retval]*/ long *hRes);
    STDMETHOD(SaveToELANFlexTime)(BSTR ElanPath, /*[out, retval]*/ long *hRes);

protected:
    int safeArraySize;
    CKhParser pureParser;
    std::map<wchar_t, wchar_t> symbols;

    void mapFileName(BSTR fileName);
};

#endif
