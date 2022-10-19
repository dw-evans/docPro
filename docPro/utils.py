
from asyncore import write
from concurrent.futures import process
import win32com.client
import json
import os
import datetime
from pathlib import Path


class VersionInfo:
    author = 'Daniel Evans'
    last_edited = '19 Oct 2022'
    version_number = '0.0.1'



def writeTxt(text, fpath):
    with open(fpath, 'w') as f:
        f.write(text)


def initialiseWord(visible=True):
    global WD
    # Initialise Word, always visible so it doesn't need to be ended
    # via task manager if bug occurs
    print("Initialising Word...")
    WD = win32com.client.Dispatch("Word.Application")
    WD.Visible = visible
    print("Word Initialised...")


def initialiseDicts():
    global PREFIX_DICT, REGEX_DICT, URLBASE_DICT, ADDLINKS_DICT, REFERENCES_DICT

    PREFIX_DICT = {}
    PREFIX_DICT['hc_1'] = "UI,PMCF,PMSP,CDE,CEP,CER,CIA,CMD,DHF,DPM,ETH,FAI,FMA,HCP,ITA,ITI,ITF,ITO,KDN,MCR,MSD,MSP,PAC,PCB,PDP,PHA,PRR,PSW,PSX,PTM,PTX,PVD,QIC,REG,SOP,SUI,SYM,TMP,TRM,TRS,CI,CP,DD,DI,DP,DR,DS,ES,FI,FM,GD,HA,HR,IU,JI,KP,MF,MT,PF,PP,PV,QR,RR,SA,SP,SW,TB,TM,TP,TR,TS,US,BE,HHE,RAM".split(",")
    PREFIX_DICT['hc_2'] = "LM,AS".split(",")
    PREFIX_DICT['ecn'] =  "ECN,EPC".split(",")
    PREFIX_DICT['dwg'] =  "DWG".split(",")

    # Can add multiple search criteria to pick up bad formatting, e.g. DWG and DRW
    REGEX_DICT = {}
    REGEX_DICT['hc_1'] = ['[A-Z]{1,4}[-^~]([0-9]{1,14})>']
    REGEX_DICT['hc_2'] = ['AS[-^~]([A-Z0-9]{1,14})>', 'LM[-^~]([A-Z0-9]{1,14})>']
    # correct ECN/EPC to search for ECNXXXXX, correct hyphen remover for this too
    REGEX_DICT['ecn']  = ['ECN[-^~]([0-9]{4,6})>', 'ECN([0-9]{4,6})>', 'EPC[-^~]([0-9]{4,6})>', 'EPC([0-9]{4,6})>']  # ['<(ECN)*([0-9]{4,6})>','<(EPC)*([0-9]{4,6})>']
    REGEX_DICT['dwg']  = ['DWG[-^~][0-9]{1,14}', 'DRW[-^~][0-9]{1,14}']

    URLBASE_DICT = {}
    URLBASE_DICT['hc_1'] =  "http://www-hcr:90/_intra_dr/hcdocs/released/{}.pdf"
    URLBASE_DICT['hc_2'] =  URLBASE_DICT['hc_1']
    URLBASE_DICT['ecn'] =   "https://search.fphcare.com/api/documents/ecn/{}.pdf"
    URLBASE_DICT['dwg'] =   None  # to throw error

    # this is redundant now
    ADDLINKS_DICT = {}
    ADDLINKS_DICT['hc_1'] = True
    ADDLINKS_DICT['hc_2'] = ADDLINKS_DICT['hc_1']
    ADDLINKS_DICT['ecn'] = ADDLINKS_DICT['hc_1']
    ADDLINKS_DICT['dwg'] = False  

    REFERENCES_DICT = {}

initialiseDicts()


def initialiseMetaFolder():
    dir_target = DOC.Path
    dir_metadata = os.path.join(dir_target, 'meta')
    if not os.path.exists(dir_metadata):
        os.makedirs(dir_metadata)
    return dir_metadata


def getMetaData():
    # gathers metadata for json export
    runtime_dict = {}
    tmp1 = datetime.datetime.now()
    runtime_dict['processed_when'] = {
        "date": '{}-{}-{}'.format(tmp1.year, tmp1.month, tmp1.day),
        "time": '{}:{}:{}'.format(tmp1.hour, tmp1.minute, tmp1.second),
    }
    # https://docs.microsoft.com/en-us/office/vba/api/word.wdbuiltinproperty
    wd_file_properties_dict = {}
    tmp = DOC.BuiltInDocumentProperties
    wd_file_properties_dict['file_name'] = DOC.Name
    tmp1 = tmp[11].Value

    wd_file_properties_dict['saved_when'] = {
        "date": '{}-{}-{}'.format(tmp1.year, tmp1.month, tmp1.day),
        "time": '{}:{}:{}'.format(tmp1.hour, tmp1.minute, tmp1.second),
        }
    wd_file_properties_dict['saved_by'] = tmp[6].Value
    wd_file_properties_dict['title'] = tmp[0].Value
    wd_file_properties_dict['size_bytes'] = tmp[21].Value
    wd_file_properties_dict['pages'] = tmp[13].Value
    wd_file_properties_dict['words'] = tmp[14].Value
    wd_file_properties_dict['characters'] = tmp[15].Value

    metadata_dict = {}
    metadata_dict['runtime_dict'] = runtime_dict
    metadata_dict['wd_file_properties_dict'] = wd_file_properties_dict
    metadata_dict['reference_list_dict'] = REFERENCES_DICT

    return metadata_dict


def getReferencesAsList():
    # generates list of references for txt file
    reference_list = []
    for key in REFERENCES_DICT:
        for string in REFERENCES_DICT[key]:
            reference_list.append(string)
    return reference_list


def referenceType(chars):
    # takes input of chars and attempts to match chars to items in keys of prefix_list_dict
    # returns key if successful or None if not found
    refType = None
    for key in PREFIX_DICT:
        if chars in PREFIX_DICT[key]:
            refType = key
    return refType


def processFile(wdDocument, addHyperlinks=True):
    global DOC

    DOC = wdDocument
    doc = DOC

    NB_HYPH = '\x1e'  # this is word's representation of a nonbreaking hyphen

    search_type_list = [key for key in REGEX_DICT]

    for search_type in search_type_list:
        reference_list = []
        print("Searching for: {}".format(search_type))

        for search_str in REGEX_DICT[search_type]:
            print("Searching for: {1} within {0}".format(search_type, search_str))
            rng = doc.Range()

            rng.Find.Text = search_str
            rng.Find.MatchWildcards = True
            rng.Find.Forward = False  # search direction
            rng.Find.Wrap = 0  # see wdFindStop enumeration

            while rng.Find.Execute():
                str_rng = rng.Text

                # create url strings and prefixes
                if search_type in ['hc_1','hc_2']:
                    str_db = str_rng.replace(NB_HYPH, "-")
                    str_url = str_db
                    str_pre = str_db.split('-')[0]
                if search_type in ['ecn']:
                    # I see some inefficiencies in here
                    # format ECN-XXXXX and ECNXXXXX
                    # ECN/EPC always 3 character prefix
                    str_pre = search_str[0:3]
                    if '[-^~]' in search_str:  
                        str_db = str_rng.replace(NB_HYPH, "-")
                        str_suff = str_db.split("-")[1]
                        str_url = str_db.replace("-","")
                    else:
                        # If no hyphen  in ECN...
                        str_url = str_rng
                        str_suff = str_rng[3:]
                    str_db = str_pre + "-" + str_suff
                if search_type in ['dwg']:
                    str_db = str_rng.replace(NB_HYPH, "-")
                    str_suff = str_db.split('-')[-1]
                    str_url = str_db  # this should never be used as no hyperlinking for dwg
                    str_pre = 'DWG'  # set dwg prefix to always be 'DWG'
                    
                isValid = str_pre in PREFIX_DICT[search_type]

                if isValid:
                    reference_list.append(str_db)
                    # if the user wants to add hyperlinks, and if the link can be made
                    if addHyperlinks and ADDLINKS_DICT[search_type]:
                        url_address = URLBASE_DICT[search_type].format(str_url)
                        # adds hyperlink, preserves original rng.Text
                        doc.Hyperlinks.Add(Anchor=rng, Address=url_address, 
                                ScreenTip="", TextToDisplay=rng.Text)

                rng.Collapse(1)
            
            reference_list = list(set(reference_list))  # remove duplicates
            reference_list.sort()  # sort
            REFERENCES_DICT[search_type] = reference_list


def setDOC(document):
    global DOC
    DOC = document


def writeMetaData():

    print("Fetching Metadata...")
    metadata_dict = getMetaData()
    reference_list = getReferencesAsList()

    dir_metadata = initialiseMetaFolder()

    references_txt = [ref+"\n" for ref in reference_list][0]
    references_json = json.dumps(metadata_dict, indent=4)

    print("Writing Metadata...")

    writeTxt(references_txt, (Path(dir_metadata) / "references.txt"))
    writeTxt(references_json, (Path(dir_metadata) / "meta.json"))

def run():
    # commandline version
    initialiseWord()
    print("\nActivate your target document.")
    # input("Continue (Enter)")
    user_input = 'n'
    while user_input != 'y':
        if user_input == 'q':
            quit()
        setDOC(WD.ActiveDocument)
        print("\nActive Document is {}".format(DOC.Name))
        user_input = input('Continue (y), Retry (Enter), Quit (q): '.format(DOC.Name)).lower()

    processFile(DOC, addHyperlinks=True)
    writeMetaData()

