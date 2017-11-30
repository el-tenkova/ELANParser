// CKhParserELAN.cpp : Implementation of ELAN's part

#include "stdafx.h"
#include <locale.h>
#include <fstream>
#include <codecvt>

#include "KhakParser_i.h"
#include "CKhParser.h"

HRESULT CKhParser::SaveToELAN(BSTR ElanPath, /*[out, retval]*/ long *hRes)
{
    *hRes = S_OK;
    std::locale loc = std::locale(std::locale("C"), new std::codecvt_utf8<wchar_t, 0x10ffff, std::generate_header>());
    std::wofstream elan(ElanPath, std::wofstream::binary);
    if (elan.is_open()) {
        elan.imbue(loc);

        writeHeader(elan);

        writeTimes(elan);

        _ULonglong id = 0;
        if (names.size() == 0)
            names.insert(std::pair<std::wstring, int>(std::wstring(L""), 1));
        for (size_t i = 1; i <= names.size(); i++) {
            std::map<std::wstring, int>::iterator nit = names.begin();
            for (; nit != names.end(); ++nit) {
                if (i == nit->second) {
                    cur_name = nit->first;
                    break;
                }
            }
            id = writeKhakSent(elan, id, size);

            id = writeWords(elan, id);

            _ULonglong refid = id;

            id = writeKhakHoms(elan, id);

            id = writeLemma(elan, id, refid);

            id = writePartOfSpeech(elan, id, refid);

            refid = id;
            id = writeKhakMorphems(elan, id);

            id = writeRusMorphems(elan, id, refid);

            id = writeEngMorphems(elan, id, refid);

            //id = writeRusHoms(elan, id, refid);

            id = writeRusSent(elan, id);
        }
        writeTail(elan);

        elan.close();
    }
    return S_OK;
}

HRESULT CKhParser::SaveToELANFlex(BSTR ElanPath, /*[out, retval]*/ long *hRes)
{
    *hRes = S_OK;
    std::locale loc = std::locale(std::locale("C"), new std::codecvt_utf8<wchar_t, 0x10ffff, std::generate_header>());
    std::wofstream elan(ElanPath, std::wofstream::binary);
    if (elan.is_open()) {
        elan.imbue(loc);

        writeHeader(elan);

        writeSentOnlyTimes(elan);

        _ULonglong id = 0;
        if (names.size() == 0)
            names.insert(std::pair<std::wstring, int>(std::wstring(L""), 1));
        for (size_t i = 1; i <= names.size(); i++) {
            std::map<std::wstring, int>::iterator nit = names.begin();
            for (; nit != names.end(); ++nit) {
                if (i == nit->second) {
                    cur_name = nit->first;
                    break;
                }
            }
            _ULonglong refid = id;
            
            id = writeKhakSent(elan, id, simple);

            _ULonglong refidWords = id;

            id = writeWordsAsRef(elan, id, refid);

            id = writeKhakHomsAsRef(elan, id, refidWords);

            id = writeLemma(elan, id, refidWords);

            id = writePartOfSpeech(elan, id, refidWords);

/*            refid = id;
            id = writeKhakMorphems(elan, id); */

//            id = writeRusMorphems(elan, id, refidWords);

//            id = writeEngMorphems(elan, id, refid);

            id = writeRusHoms(elan, id, refid);

            id = writeEngHoms(elan, id, refid);

            id = writeRusSent(elan, id);
        }
        writeTail(elan);

        elan.close();
    }
    return S_OK;
}

void CKhParser::writeHeader(std::wofstream& ef)
{
    std::wstring header(L"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    header.append(L"<ANNOTATION_DOCUMENT AUTHOR=\"unspecified\" DATE=\"2015-04-09T12:42:10+04:00\" FORMAT=\"2.8\" VERSION=\"2.8\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:noNamespaceSchemaLocation=\"http://www.mpi.nl/tools/elan/EAFv2.8.xsd\">\n");
    header.append(L"<HEADER MEDIA_FILE=\"\" TIME_UNITS=\"milliseconds\">\n");
 //   header.append(L"<MEDIA_DESCRIPTOR MEDIA_URL=\"\" MIME_TYPE=\"audio/x-wav\"/>\n"); //to avoid Annex reject error
    //'<PROPERTY NAME="URN">urn:nl-mpi-tools-elan-eaf:6504c785-9396-4c0d-abc5-c224bfd7eb09</PROPERTY>
    //    '<PROPERTY NAME="lastUsedAnnotationId">49</PROPERTY>
    header.append(L"</HEADER>\n");
    ef.write(header.c_str(), header.length());
}

void CKhParser::writeTail(std::wofstream& ef)
{
    std::wstring tail(L"<LINGUISTIC_TYPE GRAPHIC_REFERENCES = \"false\" LINGUISTIC_TYPE_ID=\"paragraph\" TIME_ALIGNABLE=\"true\" />\n");
    tail.append(L"<LINGUISTIC_TYPE CONSTRAINTS=\"Symbolic_Association\" GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"association\" TIME_ALIGNABLE=\"false\" />\n");
    tail.append(L"<LINGUISTIC_TYPE CONSTRAINTS=\"Symbolic_Subdivision\" GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"subdivision\" TIME_ALIGNABLE=\"false\" />\n");
    tail.append(L"<CONSTRAINT DESCRIPTION=\"Time subdivision of parent annotation's time interval, no time gaps allowed within this interval\" STEREOTYPE=\"Time_Subdivision\" />\n");
    tail.append(L"<CONSTRAINT DESCRIPTION=\"Symbolic subdivision of a parent annotation. Annotations refering to the same parent are ordered\" STEREOTYPE=\"Symbolic_Subdivision\" />\n");
    tail.append(L"<CONSTRAINT DESCRIPTION=\"1-1 association with a parent annotation\" STEREOTYPE=\"Symbolic_Association\" />\n");
    tail.append(L"<CONSTRAINT DESCRIPTION=\"Time alignable annotations within the parent annotation's time interval, gaps are allowed\" STEREOTYPE=\"Included_In\" />\n");
/*    std::wstring tail(L"<LINGUISTIC_TYPE GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"default\" TIME_ALIGNABLE=\"true\"/>\n");
    tail.append(L"<LINGUISTIC_TYPE GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"ruswords\" TIME_ALIGNABLE=\"false\"/>\n");
    tail.append(L"<LINGUISTIC_TYPE GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"engwords\" TIME_ALIGNABLE=\"false\"/>\n");
    tail.append(L"<LINGUISTIC_TYPE GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"lemmata\" TIME_ALIGNABLE=\"false\"/>\n");
    tail.append(L"<LINGUISTIC_TYPE GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"partofspeech\" TIME_ALIGNABLE=\"false\"/>\n");
    tail.append(L"<LINGUISTIC_TYPE CONSTRAINTS=\"Symbolic_Association\" GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"association\" TIME_ALIGNABLE=\"false\"/>\n");
    tail.append(L"<LINGUISTIC_TYPE CONSTRAINTS=\"Symbolic_Subdivision\" GRAPHIC_REFERENCES=\"false\" LINGUISTIC_TYPE_ID=\"subdivision\" TIME_ALIGNABLE=\"false\"/>\n"); */
    tail.append(L"<LOCALE COUNTRY_CODE=\"RU\" LANGUAGE_CODE=\"ru\"/>\n");
/*    tail.append(L"<CONSTRAINT DESCRIPTION=\"Time subdivision of parent annotation's time interval, no time gaps allowed within this interval\" STEREOTYPE=\"Time_Subdivision\"/>\n");
    tail.append(L"<CONSTRAINT DESCRIPTION=\"Symbolic subdivision of a parent annotation. Annotations refering to the same parent are ordered\" STEREOTYPE=\"Symbolic_Subdivision\"/>\n");
    tail.append(L"<CONSTRAINT DESCRIPTION=\"1-1 association with a parent annotation\" STEREOTYPE=\"Symbolic_Association\"/>\n");
    tail.append(L"<CONSTRAINT DESCRIPTION=\"Time alignable annotations within the parent annotation's time interval, gaps are allowed\" STEREOTYPE=\"Included_In\"/>\n"); */
    tail.append(L"</ANNOTATION_DOCUMENT>\n");
    ef.write(tail.c_str(), tail.length());
}

void CKhParser::writeTimeSlot(std::wofstream& ef, const _ULonglong& idx, const _ULonglong& begin)
{
    std::wstring str(L"<TIME_SLOT TIME_SLOT_ID=\"ts");
    str.append(std::to_wstring(idx));
    str.append(L"\" TIME_VALUE=\"");
    str.append(std::to_wstring(begin));
    str.append(L"\" />\n");
    ef.write(str.c_str(), str.length());
}

void CKhParser::writeAnno(std::wofstream& ef, const std::wstring& sent, const _ULonglong& idx, const _ULonglong& time1, const _ULonglong& time2)
{
//    text = Replace(text, "&H001f", "")
    std::wstring anno(L"<ANNOTATION>\n");
    anno.append(L"<ALIGNABLE_ANNOTATION ANNOTATION_ID=\"a");
    anno.append(std::to_wstring(idx));
    anno.append(L"\" TIME_SLOT_REF1=\"ts");
    anno.append(std::to_wstring(time1));
    anno.append(L"\" TIME_SLOT_REF2=\"ts");
    anno.append(std::to_wstring(time2));
    anno.append(L"\" >\n");
    anno.append(L"<ANNOTATION_VALUE>");
    anno.append(sent);
    anno.append(L"</ANNOTATION_VALUE>\n");
    anno.append(L"</ALIGNABLE_ANNOTATION>\n");
    anno.append(L"</ANNOTATION>\n");
    ef.write(anno.c_str(), anno.length());
}

void CKhParser::writeRefAnno(std::wofstream& ef, const std::wstring& sent, const _ULonglong& idx, const _ULonglong& refid, const _ULonglong& previous)
{
    std::wstring anno(L"<ANNOTATION>\n");
    anno.append(L"<REF_ANNOTATION ANNOTATION_ID=\"a");
    anno.append(std::to_wstring(idx));
    anno.append(L"\" ANNOTATION_REF=\"a");
    anno.append(std::to_wstring(refid));
    if (previous != (_ULonglong)-1) {
        anno.append(L"\" PREVIOUS_ANNOTATION=\"a");
        anno.append(std::to_wstring(previous));
    }
    anno.append(L"\" >\n");
    anno.append(L"<ANNOTATION_VALUE>");
    anno.append(sent);
    anno.append(L"</ANNOTATION_VALUE>\n");
    anno.append(L"</REF_ANNOTATION>\n");
    anno.append(L"</ANNOTATION>\n");
    ef.write(anno.c_str(), anno.length());
}

void CKhParser::writeTimes(std::wofstream& ef)
{
    SentVct::iterator it = sentences.begin();
    _ULonglong idx = 1;
    _ULonglong begin = 0;
    std::wstring h(L"<TIME_ORDER>\n");
    ef.write(h.c_str(), h.length());
    for (; it != sentences.end(); ++it) {
        for (size_t i = 0; i < it->size; i++) {
            writeTimeSlot(ef, idx, begin);
            begin += 500;
            idx += 1;
        } 
    }
    // временно
    for (size_t i = 0; i < 20; i++) {
        writeTimeSlot(ef, idx, begin);
        begin += 500;
        idx += 1;
    }
  //  writeTimeSlot(ef, idx, begin);
  //  writeTimeSlot(ef, idx + 1, begin + 500);
    h.erase();
    h = L"</TIME_ORDER>\n";
    ef.write(h.c_str(), h.length());
}

void CKhParser::writeSentOnlyTimes(std::wofstream& ef)
{
    SentVct::iterator it = sentences.begin();
    _ULonglong idx = 1;
    _ULonglong begin = 0;
    std::wstring h(L"<TIME_ORDER>\n");
    ef.write(h.c_str(), h.length());
    for (; it != sentences.end(); ++it) {
        writeTimeSlot(ef, idx, begin);
        begin += 5000;
        idx += 1;
    }
    // временно
    for (size_t i = 0; i < 2; i++) {
        writeTimeSlot(ef, idx, begin);
        begin += 5000;
        idx += 1;
    }
    //  writeTimeSlot(ef, idx, begin);
    //  writeTimeSlot(ef, idx + 1, begin + 500);
    h.erase();
    h = L"</TIME_ORDER>\n";
    ef.write(h.c_str(), h.length());
}

void CKhParser::writeTierHeader(std::wofstream& ef, const std::wstring& name, const std::wstring& type, const std::wstring& parent, const std::wstring& participant)
{
    std::wstring tier(L"<TIER DEFAULT_LOCALE=\"ru\" LINGUISTIC_TYPE_REF=\"");
    tier.append(type);
    tier.append(L"\" TIER_ID=\"");
    tier.append(name);
    tier.append(L"\"");
    if (parent.length() > 0) {
        tier.append(L" PARENT_REF=\"");
        tier.append(parent);
        tier.append(L"\"");
    }
    if (participant.length() > 0) {
        tier.append(L" PARTICIPANT=\"");
        tier.append(participant);
        tier.append(L"\"");
    }
    tier.append(L" >\n");
    ef.write(tier.c_str(), tier.length());
}

void CKhParser::writeTierTail(std::wofstream& ef)
{
    std::wstring tier(L"</TIER>\n");
    ef.write(tier.c_str(), tier.length());
}

_ULonglong CKhParser::writeKhakSent(std::wofstream& ef, _ULonglong& id, const stepType timeStep)
{
    _ULonglong time1 = 1;
    _ULonglong time2 = 1;
    std::wstring lvlName = lvlNames[Kh_Sent];
    appendName(lvlName, std::wstring());
    writeTierHeader(ef, lvlName, L"paragraph", L"", cur_name);

    SentVct::iterator it = sentences.begin();
    for (; it != sentences.end(); ++it) {
        if (timeStep == simple)
            time2 = time1 + 1;
        else
            time2 = time1 + it->size;
        if (it->informant == cur_name) {
            writeAnno(ef, it->khak_sent, id, time1, time2);
            it->id = id;
            id++;
            it->time1 = time1;
            it->time2 = time2;
        }
        time1 = time2;
    }
    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeRusSent(std::wofstream& ef, _ULonglong& id)
{
    std::wstring lvlName = lvlNames[Rus_Sent];
    std::wstring khLvlName = lvlNames[Kh_Sent];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"association", khLvlName, cur_name);

    SentVct::iterator it = sentences.begin();
    _ULonglong previous = (_ULonglong)(-1);
    for (; it != sentences.end(); ++it) {
        if (it->informant == cur_name) {
            writeRefAnno(ef, it->rus_sent, id, it->id, previous);
            id++;
        }
    }
    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeWords(std::wofstream& ef, _ULonglong& id)
{
    _ULonglong time1 = 1;
    _ULonglong time2 = 1;
    std::wstring lvlName = lvlNames[Kh_Words];
    std::wstring khLvlName = lvlNames[Kh_Sent];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"subdivision", khLvlName, cur_name);

    SentVct::iterator it = sentences.begin();
    for (; it != sentences.end(); ++it) {
        if (it->informant == cur_name) {
            time1 = it->time1;
            std::vector<std::wstring>::iterator wt = it->words.begin();
            for (; wt != it->words.end(); ++wt) {
                HomMap::iterator ct = cache.find((it->keys.find(*wt))->second);
                if (ct != cache.end()) {
                    //time2 = time1 + ct->second.homVct.size();
                    time2 = time1 + getMorphemsCount(ct);
                }
                else
                    time2 = time1 + 1;
                writeAnno(ef, (*wt), id, time1, time2);
                time1 = time2;
                id++;
            }
        }
    }
    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeWordsAsRef(std::wofstream& ef, _ULonglong& id, _ULonglong refid)
{

    std::wstring lvlName = lvlNames[Kh_Words];
    std::wstring khLvlName = lvlNames[Kh_Sent];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"subdivision", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, words);

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeKhakHoms(std::wofstream& ef, _ULonglong& id)
{
    _ULonglong time1 = 1;
    _ULonglong time2 = 1;
    std::wstring lvlName = lvlNames[Kh_Homonyms];
    std::wstring khLvlName = lvlNames[Kh_Words];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"default", khLvlName, cur_name);

    SentVct::iterator it = sentences.begin();
    for (; it != sentences.end(); ++it) {
        if (it->informant == cur_name) {
            it->firstHomId = id;
            time1 = it->time1;
            time2 = time1;
            std::vector<std::wstring>::iterator wt = it->words.begin();
            for (; wt != it->words.end(); ++wt) {
                HomMap::iterator ct = cache.find((it->keys.find(*wt))->second);
                if (ct != cache.end()) {
                    HomVct::iterator vt = ct->second.homVct.begin();
                    for (; vt != ct->second.homVct.end(); ++vt) {
                        //time2++;
                        time2 += vt->morphems.size();
                        writeAnno(ef, vt->khak, id, time1, time2);
                        id++;
                        time1 = time2;
                    }
                }
                else {
                    time1++;
                    time2 = time1;
                }
            }
        }
    }

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeKhakHomsAsRef(std::wofstream& ef, _ULonglong& id, _ULonglong refid)
{

    std::wstring lvlName = lvlNames[Kh_Homonyms];
    std::wstring khLvlName = lvlNames[Kh_Words];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"subdivision", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, homs);

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeKhakMorphems(std::wofstream& ef, _ULonglong& id)
{
    _ULonglong time1 = 1;
    _ULonglong time2 = 1;
    std::wstring lvlName = lvlNames[Kh_Morphems];
    std::wstring khLvlName = lvlNames[Kh_Words];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"default", khLvlName, cur_name);

    SentVct::iterator it = sentences.begin();
    for (; it != sentences.end(); ++it) {
        if (it->informant == cur_name) {
            it->firstHomId = id;
            time1 = it->time1;
            time2 = time1;
            std::vector<std::wstring>::iterator wt = it->words.begin();
            for (; wt != it->words.end(); ++wt) {
                HomMap::iterator ct = cache.find((it->keys.find(*wt))->second);
                if (ct != cache.end()) {
                    HomVct::iterator vt = ct->second.homVct.begin();
                    for (; vt != ct->second.homVct.end(); ++vt) {
                        for (auto mit = vt->morphems.begin(); mit != vt->morphems.end(); ++mit) {
                            time2++;
                            writeAnno(ef, *(mit), id, time1, time2);
                            id++;
                            time1 = time2;
                        }
                    }
                }
                else {
                    time1++;
                    time2 = time1;
                }
            }
        }
    }
    writeTierHeader(ef, lvlName, L"paragraph", khLvlName, cur_name);
    return id;
}

_ULonglong CKhParser::writeRefTier(std::wofstream& ef, _ULonglong& id, _ULonglong _refid, const reftype& type)
{
    SentVct::iterator it = sentences.begin();
    _ULonglong sentid = _refid;
    _ULonglong wordid = _refid;
    for (; it != sentences.end(); ++it) {
        if (it->informant == cur_name) {
            if (type == homs) {
                it->firstHomId = id;
            }
            _ULonglong refid = it->firstHomId;
            std::vector<std::wstring>::iterator wt = it->words.begin();
            _ULonglong previous = (_ULonglong)(-1);
            for (; wt != it->words.end(); ++wt) {
                if (type == words) {
                    writeRefAnno(ef, (*wt), id, sentid, previous);
                    previous = id;
                    id++;
                    continue;
                }
                HomMap::iterator ct = cache.find((it->keys.find(*wt))->second);
                if (ct != cache.end()) {
                    HomVct::iterator vt = ct->second.homVct.begin();
                    for (; vt != ct->second.homVct.end(); ++vt) {
                        switch (type)
                        {
                            case homs: 
                            {
                                writeRefAnno(ef, vt->khak, id, wordid, previous); 
                                previous = id;
                                id++;
                                break;
                            }
                            case m_eng:
                            case m_rus:
                            {
                                for (auto mit = vt->r_morphems.begin(); mit != vt->r_morphems.end(); ++mit) {
                                    writeRefAnno(ef, (*mit), id, refid, previous);
                                    id++;
                                    refid++;
                                }
                                break;
                            }
                            case rus: writeRefAnno(ef, vt->rus, id, refid, previous); id++; refid++; break;
                            case lemma: writeRefAnno(ef, vt->lemma, id, refid, previous); id++; refid++; break;
                            case partofspeech: writeRefAnno(ef, vt->pos, id, refid, previous); id++; refid++; break;
                            default: break;
                        }
                    }
                }
                wordid++;
                previous = (_ULonglong)(-1);
            }
            sentid++;
        }
    }
    return id;
}

_ULonglong CKhParser::writeRusMorphems(std::wofstream& ef, _ULonglong& id, _ULonglong& refid)
{
    std::wstring lvlName = lvlNames[Rus_Morphems];
    std::wstring khLvlName = lvlNames[Kh_Morphems];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"ruswords", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, m_rus);

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeEngMorphems(std::wofstream& ef, _ULonglong& id, _ULonglong& refid)
{
    std::wstring lvlName = lvlNames[Eng_Morphems];
    std::wstring khLvlName = lvlNames[Kh_Morphems];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"engwords", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, m_eng);

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeRusHoms(std::wofstream& ef, _ULonglong& id, _ULonglong& refid)
{
    std::wstring lvlName = lvlNames[Rus_Homonyms];
    std::wstring khLvlName = lvlNames[Kh_Homonyms];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"association", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, rus);

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeEngHoms(std::wofstream& ef, _ULonglong& id, _ULonglong& refid)
{
    std::wstring lvlName = lvlNames[Eng_Homonyms];
    std::wstring khLvlName = lvlNames[Kh_Homonyms];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"association", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, rus);

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writeLemma(std::wofstream& ef, _ULonglong& id, _ULonglong& refid)
{
    std::wstring lvlName = lvlNames[Kh_Lemma];
    std::wstring khLvlName = lvlNames[Kh_Homonyms];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"association", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, lemma);

    writeTierTail(ef);
    return id;
}

_ULonglong CKhParser::writePartOfSpeech(std::wofstream& ef, _ULonglong& id, _ULonglong& refid)
{
    std::wstring lvlName = lvlNames[Kh_PartOfSpeech];
    std::wstring khLvlName = lvlNames[Kh_Homonyms];
    appendName(lvlName, khLvlName);

    writeTierHeader(ef, lvlName, L"association", khLvlName, cur_name);

    id = writeRefTier(ef, id, refid, partofspeech);

    writeTierTail(ef);
    return id;
}

void CKhParser::appendName(std::wstring& lvlName, std::wstring& refLvlName)
{
    if (cur_name.length() != 0) {
/*        lvlName.append(L"@");
        lvlName.append(cur_name);
        if (refLvlName.length() != 0) {
            refLvlName.append(L"@");
            refLvlName.append(cur_name);
        } */
        lvlName = cur_name + lvlName;
        if (refLvlName.length() != 0) {
            refLvlName = cur_name + refLvlName;
        }
    }
}
