#ifndef __KHHTTP_H_
#define __KHHTTP_H_

#ifdef _USE_MSXML2

#import "msxml6.dll"
using namespace MSXML2;

#include <string>

class CKhHttpWrapper
{
protected:
    std::string www;
    std::string request;
    IXMLHTTPRequestPtr pIXMLHTTPRequest;

public:
    CKhHttpWrapper(long* hRes, const char* www);
    ~CKhHttpWrapper(void);

public:
    long Open(const wchar_t* toPost);
    long Send(void);
    size_t Status(void);
    const wchar_t* ResponseText(void);
};
#endif
#endif // __KHHTTP_H_
