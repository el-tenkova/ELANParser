
#ifndef __KHPARSERPURE_H_
#define __KHPARSERPURE_H_

#include <locale.h>

#include <string>
#include <vector>
#include <map>

#ifdef _WINDOWS
    #include "resource.h"       // main symbols
#endif

#ifdef _USE_MSXML2
#include "CKhHttp.h"
#else
#include "CKhHttpL.h"
#endif

enum reftype
{
    words = 0,
    homs,
    rus,
    lemma,
    partofspeech,
    m_rus,
    m_eng,
};

struct hom
{
    std::wstring khak;
    std::wstring rus;
    std::wstring lemma;
    std::wstring pos;
    std::vector<std::wstring> morphems;
    std::vector<std::wstring> r_morphems;
};

typedef std::vector<hom> HomVct;

struct cacheItem
{
    int count;
    int homVctSize; // считаем для каждого омонима число морфем
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
    std::wstring informant; // имя участника разговора
    unsigned long long id; //индекс хакасского предложения
    unsigned long long firstHomId; // индекс первого омонима (для соотв. русских ref)
    unsigned long long time1; //начало разговора
    unsigned long long time2; //конец разговора
};

typedef std::vector<sent> SentVct;

class CKhParser
{
public:
    CKhParser();/*      tmpResp(L"Parsing: тустарда\n\
FOUND STEM: ----------Pl---Loc---- ----------тар---да----\n\
n тус ‛соль’ \n\
FOUND STEM: ----------Pl---Loc---- ----------тар---да----\n\
n тус ‛время’ \n\
FOUND STEM: ----------Pl---Loc---- ----------тар---да----\n\
n тус ‛напротив’ \n\
3  wordforms generated in 3571 μs.") */

    long Init(const std::wstring& www, const std::string& dict, const std::string& notfound);
    long Terminate(void);
    long DoParse(const std::wstring& InputWord);
    long AddKhakSent(const std::wstring& InputSent);
    long AddRusSent(const std::wstring& InputSent);
    long AddKhakSent2(const std::wstring& Name, const std::wstring& InputSent);
    long SaveToELAN(const std::string& ElanPath);
    long SaveToELANFlex(const std::string& ElanPath);
    int GetHomonymNum(void);
    long GetNextHomonym(std::wstring& hom);


protected:
    enum stepType {
        simple = 0,
        size,
    };
    CKhHttpWrapper* pIXMLHTTPRequest;

    std::wstring request;
    std::vector<hom> homonyms;
    int currHom;
    HomMap cache;
    SentVct sentences;
    std::map<std::wstring, int> empty;
    std::map<short, short> repl;
    std::map<std::wstring, int> names;
    std::locale russian;

#ifdef _WINDOWS
    _locale_t locinfo;
#endif

    std::string dict;
    std::string notfound;
    std::wstring cur_name; //текущий пользователь при записи слоя

    std::map<std::wstring, std::wstring> lvlNames;
    static std::wstring Kh_Sent;
    static std::wstring Kh_Words;
    static std::wstring Kh_Homonyms;
    static std::wstring Kh_Lemma;
    static std::wstring Kh_PartOfSpeech;
    static std::wstring Kh_Morphems;
    static std::wstring Rus_Morphems;
    static std::wstring Eng_Morphems;
    static std::wstring Rus_Homonyms;
    static std::wstring Eng_Homonyms;
    static std::wstring Rus_Sent;

    long normWord(const std::wstring& inputWord, std::wstring& normWord);
    long fillHomonyms(const std::wstring& response);
    wchar_t* getDetails(const std::wstring& str, wchar_t endCh);
    wchar_t* getSubstr(const wchar_t* str, wchar_t endCh);
    long saveResults(void);
    void addToSentSize(int value);
    void addWordsToSent(const std::wstring& word, const std::wstring& normWord);
    void getMorphems(const std::wstring& affixes, std::vector<std::wstring>& morphems, const wchar_t& delim);
    size_t getMorphemsCount(HomMap::iterator& ct);
    wchar_t* removeSymbols(wchar_t* input);

    // ELAN part
    void writeHeader(std::wofstream& ef);
    void writeTail(std::wofstream& ef);
    void writeTimeSlot(std::wofstream& ef, const unsigned long long& idx, const unsigned long long& begin);
    void writeTierHeader(std::wofstream& ef, const std::wstring& name, const std::wstring& type, const std::wstring& parent, const std::wstring& participant);
    void writeTierTail(std::wofstream& ef);
    void writeAnno(std::wofstream& ef, const std::wstring& sent, const unsigned long long& idx, const unsigned long long& time1, const unsigned long long& time2);
    void writeRefAnno(std::wofstream& ef, const std::wstring& sent, const unsigned long long& idx, const unsigned long long& refid, const unsigned long long& previous);

    void writeTimes(std::wofstream& ef);
    void writeSentOnlyTimes(std::wofstream& ef);
    unsigned long long writeKhakSent(std::wofstream& ef, unsigned long long& id, const stepType timeStep);
    unsigned long long writeRusSent(std::wofstream& ef, unsigned long long& id);
    unsigned long long writeWords(std::wofstream& ef, unsigned long long& id);
    unsigned long long writeKhakHoms(std::wofstream& ef, unsigned long long& id);
    unsigned long long writeWordsAsRef(std::wofstream& ef, unsigned long long& id, const unsigned long long refid);
    unsigned long long writeKhakHomsAsRef(std::wofstream& ef, unsigned long long& id, const unsigned long long refid);
    unsigned long long writeRefTier(std::wofstream& ef, unsigned long long& id, unsigned long long refid, const reftype& type);
    unsigned long long writeRusHoms(std::wofstream& ef, unsigned long long& id, unsigned long long& refid);
    unsigned long long writeEngHoms(std::wofstream& ef, unsigned long long& id, unsigned long long& refid);
    unsigned long long writeLemma(std::wofstream& ef, unsigned long long& id, unsigned long long& refid);
    unsigned long long writePartOfSpeech(std::wofstream& ef, unsigned long long& id, unsigned long long& refid);
    unsigned long long writeKhakMorphems(std::wofstream& ef, unsigned long long& id);
    unsigned long long writeRusMorphems(std::wofstream& ef, unsigned long long& id, unsigned long long& refid);
    unsigned long long writeEngMorphems(std::wofstream& ef, unsigned long long& id, unsigned long long& refid);
    void appendName(std::wstring& lvlName, std::wstring& refLvlName);
};

#endif //__KHPARSERPURE_H_
