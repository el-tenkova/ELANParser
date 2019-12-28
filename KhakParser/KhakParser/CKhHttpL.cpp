#include "stdafx.h"
#include "CKhHttpL.h"

#include <string>
#include <locale>
#include <codecvt>
#include <iostream>

CKhHttpWrapper::CKhHttpWrapper(long* hRes, const char* _www)
:russian(std::locale(RUS_LOCALE), new std::codecvt_utf8<wchar_t, 0x10ffff, std::consume_header>())
,statusCode(0)
{
    *hRes = 0;
    www.append(_www);
}
CKhHttpWrapper::~CKhHttpWrapper(void)
{
}
long CKhHttpWrapper::Open(const wchar_t* toPost)
{
    std::wstring str(toPost);
    std::wstring_convert<std::codecvt_utf8<wchar_t>> myconv;
    request = std::string(myconv.to_bytes(str));
    return 0;
}
long CKhHttpWrapper::Send(void)
{
    boost::asio::io_service io_service;
    // Get a list of endpoints corresponding to the server name.
    tcp::resolver resolver(io_service);
    tcp::resolver::query query(www, "http");
    tcp::resolver::iterator endpoint_iterator = resolver.resolve(query);

    // Try each endpoint until we successfully establish a connection.
    tcp::socket socket(io_service);
    boost::asio::connect(socket, endpoint_iterator);

    boost::asio::streambuf req;
    std::ostream request_stream(&req);
    request_stream.imbue(russian);
    request_stream << "GET " << request << " HTTP/1.0\r\n";
    request_stream << "Host: " << www << "\r\n";
    request_stream << "Accept: */*\r\n";
    request_stream << "Connection: close\r\n\r\n";

    // Send the request.
    boost::asio::write(socket, req);

    boost::asio::streambuf response;
    boost::asio::read_until(socket, response, "\r\n");

    std::istream response_stream(&response);
    std::string http_version;
    response_stream >> http_version;
    unsigned int status_code;
    response_stream >> status_code;
    std::string status_message;
    std::getline(response_stream, status_message);
    if (!response_stream || http_version.substr(0, 5) != "HTTP/")
    {
        std::cout << "Invalid response\n";
        return 1;
    }
    statusCode = status_code;
    if (status_code != 200)
    {
        std::cout << "Response returned with status code " << status_code << "\n";
        return 1;
    }

    // Read the response headers, which are terminated by a blank line.
    boost::asio::read_until(socket, response, "\r\n\r\n");

    // Process the response headers.
    std::string header;
    while (std::getline(response_stream, header) && header != "\r");

    // Read until EOF, writing data to output as we go.
    boost::system::error_code error;
    while (boost::asio::read(socket, response,
        boost::asio::transfer_at_least(1), error));
    responseText.clear();
    if (error == boost::asio::error::eof)
    {
        std::wstring_convert<std::codecvt_utf8<wchar_t>> myconv;
        boost::asio::streambuf::const_buffers_type bufs = response.data();
        size_t buf_size = response.size();
        std::string s(boost::asio::buffers_begin(bufs),
                      boost::asio::buffers_begin(bufs) + buf_size);
        responseText = std::wstring(myconv.from_bytes(s));
    }
    return 0;
}
size_t CKhHttpWrapper::Status(void)
{
    return statusCode;
}
const wchar_t* CKhHttpWrapper::ResponseText(void)
{
    return (responseText.c_str());
}
