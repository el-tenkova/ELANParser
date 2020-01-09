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

    CKhParser kp;
    kp.Init(L"http://khakas.altaica.ru", "", "");

    std::string input("");
    std::string output("");
    for (int i = 1; i < argc; i+=2) {
        if (strcmp(argv[i], "-i") == 0) {
            std::cout << argv[i + 1] << std::endl;
            input = argv[i + 1];
        }
        else if (strcmp(argv[i], "-o") == 0) {
            std::cout << argv[i + 1] << std::endl;
            output = argv[i + 1];
        }
    }

//    std::wifstream infile("/home/elan/text-b.txt", std::wifstream::binary);
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
            if (end == std::wstring::npos)
                continue;
            name = line.substr(beg, end - beg);
            std::wcout << name << std::endl;
            std::wstring khak(L"");
            beg = end + 1;
            end = line.find(L'|', beg);
            std::cout << beg << " " << end << std::endl;
            khak = line.substr(beg, end - beg);
            std::wcout << khak << std::endl;
            std::wstring rus(L"");
            rus = line.substr(end + 1);
            std::wcout << rus << std::endl;
            kp.AddKhakSent2(name, khak);
            kp.AddRusSent(rus);
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
//        kp.SaveToELANFlex("/home/elan/result.eaf");
        kp.SaveToELANFlex(output);
    }
    kp.Terminate();
    return 0;
}

