// dllmain.h : Declaration of module class.

class CKhakParserModule : public ATL::CAtlDllModuleT< CKhakParserModule >
{
public :
	DECLARE_LIBID(LIBID_KhakParserLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_KHAKPARSER, "{580EAA2D-7DCE-43E3-8BD2-3D697EBCF1C9}")
   // OBJECT_ENTRY_AUTO(CLSID_KhParser, CKhParser)
};

extern class CKhakParserModule _AtlModule;
