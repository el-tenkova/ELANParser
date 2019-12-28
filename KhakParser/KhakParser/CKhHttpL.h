#ifndef __KHHTTP_L_H_
#define __KHHTTP_L_H_

#include <string>
#include <boost/asio.hpp>

using boost::asio::ip::tcp;
//#define RUS_LOCALE "ru_RU.UTF-8"
#define RUS_LOCALE "Russian"

class CKhHttpWrapper
{
protected:
    std::string www;
    std::string request;
    std::locale russian;
    size_t statusCode;
    std::wstring responseText;

public:
    CKhHttpWrapper(long* hRes, const char* www);
    ~CKhHttpWrapper(void);

public:
    long Open(const wchar_t* toPost);
    long Send(void);
    size_t Status(void);
    const wchar_t* ResponseText(void);
};

#endif // __KHHTTP_L_H_

