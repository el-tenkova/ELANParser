/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/* 
 * File:   main.cpp
 * Author: el-tenkova
 *
 * Created on December 29, 2019, 10:07 PM
 */

#include <cstdlib>
#include <fstream>
#include <string>
#include <locale.h>
#include <fstream>
#include <codecvt>
#include <iostream>

#include "CKhParser.h"

using namespace std;

/*
 * 
 */
int main(int argc, char** argv)
{
    std::string input("");
    std::string output("");
    std::string host("http://khakas.altaica.ru");
    int mode = 1;
    // 1 name|khak|rus
    // 2 khak|rus
    // 3 khak
    for (int i = 1; i < argc; i+=2) {
        if (strcmp(argv[i], "-i") == 0) {
            std::cout << argv[i + 1] << std::endl;
            input = argv[i + 1];
        }
        else if (strcmp(argv[i], "-o") == 0) {
            std::cout << argv[i + 1] << std::endl;
            output = argv[i + 1];
        }
        else if (strcmp(argv[i], "-m") == 0) {
            std::cout << argv[i + 1] << std::endl;
            mode = atoi(argv[i + 1]);
        }
        else if (strcmp(argv[i], "-h") == 0) {
            std::cout << argv[i + 1] << std::endl;
            host = argv[i + 1];
        }
    }
    std::cout << "mode = " << std::to_string(mode) << std::endl;
    std::cout << "host = " << host << std::endl;
    CKhParser kp;
    kp.Init(host, "", "");

    std::wifstream infile(input, std::wifstream::binary);
    std::locale russian(std::locale(RUS_LOCALE), new std::codecvt_utf8<wchar_t, 0x10ffff, std::consume_header>());

    if (infile.is_open())
        infile.imbue(russian);
    {
        while (!infile.eof())
        {
            std::wstring line(L"");
            std::getline(infile, line);
            std::wstring name(L"");
            size_t beg = 0;
            size_t end = line.find(L'|', beg);
            if (end == std::wstring::npos && mode != 3)
                continue;
            name = line.substr(beg, end - beg);
            std::wcout << name << std::endl;
            std::wstring khak(L"");
            if (mode == 1 || mode == 2)
            {
                beg = end + 1;
                end = line.find(L'|', beg);
                std::cout << beg << " " << end << std::endl;
                khak = line.substr(beg, end - beg);
                std::wcout << khak << std::endl;
                std::wstring rus(L"");
                if (mode == 1)
                {
                    rus = line.substr(end + 1);
                    std::wcout << rus << std::endl;
                    kp.AddKhakSent2(name, khak);
                    kp.AddRusSent(rus);
                }
                else
                {
                    rus = khak;
                    khak = name;
                    kp.AddKhakSent(khak);
                    kp.AddRusSent(rus);
                }
            }
            else
            {
                khak = name;
                kp.AddKhakSent(khak);
                kp.AddRusSent(L"");
            }
            std::wstring khak1(khak);
            size_t pos = khak1.find_first_of(L".,;!:-?");
            while (pos != std::wstring::npos)
            {
                khak1[pos] = L' ';
                pos = khak1.find_first_of(L".,;!:-?", pos + 1);
            }
            std::wistringstream iss(khak1.c_str());
            std::vector<std::wstring> parts;
            std::copy(std::istream_iterator<std::wstring, wchar_t>(iss),
                      std::istream_iterator<std::wstring, wchar_t>(),
                      std::back_inserter(parts));
            for (auto pit = parts.begin(); pit != parts.end(); ++pit)
            {
                kp.DoParse(*pit);
            }
        }
        if (mode == 1)
            kp.SaveToELANFlex(output);
        else
            kp.SaveToELAN(output);
    }
    kp.Terminate();
    return 0;
}

