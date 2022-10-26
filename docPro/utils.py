"""
documentPro utils.py

Utility functions for documentPro appliaction and CLI


"""

import win32com.client
import json
import os
import datetime
from pathlib import Path


class VersionInfo:
    name = "documentPro"
    author = 'Daniel Evans'
    last_edited = '26 Oct 2022'
    version_number = '0.0.2'


class SearchObject:
    # object to store search parameters and results
    def __init__(self, prefix_list:list[str], regex_list:list[str], urlbase:str, addlinks:bool):
        self.prefix_list = prefix_list
        self.regex_list = regex_list
        self.urlbase = urlbase
        self.addlinks = addlinks
        self.references: list = []


def writeTxt(text, fpath) -> None:
    # writes str to fpath
    with open(fpath, 'w') as f:
        f.write(text)


def initialiseWord(visible=True) -> None:
    # Initialise Word, always visible so it doesn't need to be ended
    # via task manager if bug occurs
    global WD
    print("Initialising Word...")
    WD = win32com.client.Dispatch("Word.Application")
    WD.Visible = visible
    print("Word Initialised...")



def initialiseDicts():
    # initialises search objects dict
    global SEARCH_OBJECTS

    SEARCH_OBJECTS = {}
    SEARCH_OBJECTS['hc_1'] = SearchObject(
        prefix_list= "UI,PMCF,PMSP,CDE,CEP,CER,CIA,CMD,DHF,DPM,ETH,FAI,FMA,HCP,ITA,ITI,ITF,ITO,KDN,MCR,MSD,MSP,PAC,PCB,PDP,PHA,PRR,PSW,PSX,PTM,PTX,PVD,QIC,REG,SOP,SUI,SYM,TMP,TRM,TRS,CI,CP,DD,DI,DP,DR,DS,ES,FI,FM,GD,HA,HR,IU,JI,KP,MF,MT,PF,PP,PV,QR,RR,SA,SP,SW,TB,TM,TP,TR,TS,US,BE,HHE,RAM".split(","),
        regex_list=['[A-Z]{1,4}[-^~]([0-9]{1,14})>'],
        urlbase="http://www-hcr:90/_intra_dr/hcdocs/released/{}.pdf",
        addlinks=True
    )
    SEARCH_OBJECTS['hc_2'] = SearchObject(
        prefix_list="LM,AS".split(","),
        regex_list=[
            'AS[-^~]([A-Z0-9]{1,14})>', 
            'LM[-^~]([A-Z0-9]{1,14})>'
        ],
        urlbase=SEARCH_OBJECTS['hc_1'].urlbase,
        addlinks=True
    )
    SEARCH_OBJECTS['ecn'] = SearchObject(
        prefix_list="ECN,EPC".split(","),
        regex_list=[
            'ECN[-^~]([0-9]{4,6})>', 
            'ECN([0-9]{4,6})>', 
            'EPC[-^~]([0-9]{4,6})>', 
            'EPC([0-9]{4,6})>'
        ],
        urlbase="https://search.fphcare.com/api/documents/ecn/{}.pdf",
        addlinks=True
    )
    SEARCH_OBJECTS['dwg'] = SearchObject(
        prefix_list="DWG".split(","),
        regex_list=[
            'DWG[-^~][0-9]{1,14}',
            'DRW[-^~][0-9]{1,14}'
        ],
        urlbase='',
        addlinks=False
    )


def initialiseMetaFolder() -> str:
    # creates a /meta directory to write metadata to
    # if it isn't created.
    dir_target = DOC.Path
    dir_metadata = os.path.join(dir_target, 'meta')
    if not os.path.exists(dir_metadata):
        os.makedirs(dir_metadata)
    return dir_metadata


def getMetaData() -> dict:
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
    # metadata_dict['reference_list_dict'] = REFERENCES_DICT

    search_objects_dict = {}
    for key in SEARCH_OBJECTS:
        search_objects_dict[key] = SEARCH_OBJECTS[key].__dict__

    metadata_dict['search_objects'] = search_objects_dict

    return metadata_dict


def getReferencesAsList() -> list[str]:
    # generates list of references for txt file export
    reference_list = []
    # for key in REFERENCES_DICT:
    #     for string in REFERENCES_DICT[key]:
    for key in SEARCH_OBJECTS:
        for string in SEARCH_OBJECTS[key].references:
            reference_list.append(string)
            print(string)
    return reference_list


def referenceType(chars):
    # takes input of chars and attempts to match chars to items in keys of prefix_list_dict
    # returns key if successful or None if not found
    refType = None
    # for key in PREFIX_DICT:
        # if chars in PREFIX_DICT[key]:
    for key in SEARCH_OBJECTS:
        if chars in SEARCH_OBJECTS[key].prefix_list:
            refType = key
    return refType


def processFile(wdDocument, addHyperlinks=True) -> None:
    """
    Main processor function. Defines the global DOC variable which is used
    in other functions.
    Perhaps this would be better implemented with a Session class or similar
    """
    global DOC

    DOC = wdDocument

    NB_HYPH = '\x1e'  # this is word's representation of a nonbreaking hyphen

    # complete a regex search for each search type (dict keys)
    # search_type_list = [key for key in REGEX_DICT]
    for search_type in SEARCH_OBJECTS:
    # for search_type in search_type_list:
        reference_list = []  # initialise ref list for metadata write
        print("Searching for: {}".format(search_type))


        # for search_str in REGEX_DICT[search_type]:
        for search_str in SEARCH_OBJECTS[search_type].regex_list:
            print("Searching for: {1} within {0}".format(search_type, search_str))
            rng = DOC.Range()  # document range (all text content in body)
            # Setup the find method variables
            rng.Find.Text = search_str  # look for this string
            rng.Find.MatchWildcards = True  # for regex to work
            rng.Find.Forward = False  # search direction, I remember it had to go backwards
            rng.Find.Wrap = 0  # see wdFindStop enumeration

            while rng.Find.Execute():
                str_rng = rng.Text  # range shifts to found text btw

                # create url strings and prefixes
                # prefix will be checked to see if valid
                if search_type in ['hc_1','hc_2']:
                    # Regular doc handling
                    str_db = str_rng.replace(NB_HYPH, "-")
                    str_url = str_db
                    str_pre = str_db.split('-')[0]
                elif search_type in ['ecn']:
                    # ECN handling
                    str_pre = search_str[0:3]
                    if '[-^~]' in search_str:  
                        # If there is a hyphen in the regex
                        str_db = str_rng.replace(NB_HYPH, "-")
                        str_suff = str_db.split("-")[1]
                        str_url = str_db.replace("-","")
                    else:
                        # otherwise...
                        str_url = str_rng
                        str_suff = str_rng[3:]
                    str_db = str_pre + "-" + str_suff
                elif search_type in ['dwg']:
                    # DWG handling
                    str_db = str_rng.replace(NB_HYPH, "-")
                    str_suff = str_db.split('-')[-1]
                    str_url = str_db  # this should never be used as no hyperlinking for dwg
                    str_pre = 'DWG'  # set dwg prefix to always be 'DWG'
                    
                # Check validity of prefix
                # isValid = str_pre in PREFIX_DICT[search_type]
                isValid = str_pre in SEARCH_OBJECTS[search_type].prefix_list

                if isValid:
                    
                    reference_list.append(str_db)
                    # if the user wants to add hyperlinks, and if the link can be made
                    # if addHyperlinks and ADDLINKS_DICT[search_type]:
                    if addHyperlinks and SEARCH_OBJECTS[search_type].addlinks:
                        # url_address = URLBASE_DICT[search_type].format(str_url)
                        url_address = SEARCH_OBJECTS[search_type].urlbase.format(str_url)
                        # adds hyperlink, preserves original rng.Text
                        DOC.Hyperlinks.Add(Anchor=rng, Address=url_address, 
                                ScreenTip="", TextToDisplay=rng.Text)

                # collapse the range to the end of the word
                # Not 100% sure why by now, perhaps to stop infinite looping
                rng.Collapse(1)
            
            # remove duplicate occurances and sort alphabetically.
            reference_list = list(set(reference_list)) 
            reference_list.sort() 

            # write the reference list to the global results dict
            # REFERENCES_DICT[search_type] = reference_list
            SEARCH_OBJECTS[search_type].references = reference_list


def setDOC(document) -> None:
    # function to set DOC without needing a global DOC declaration
    global DOC
    DOC = document


def writeMetaData() -> None:
    # gets and writes meta data to metadata folder as json, txt

    print("Fetching Metadata...")
    metadata_dict = getMetaData()
    reference_list = getReferencesAsList()

    dir_metadata = initialiseMetaFolder()

    references_txt = ''
    for ref in reference_list:
        references_txt += ref+"\n"
    references_json = json.dumps(metadata_dict, indent=4)

    print("Writing Metadata...")

    writeTxt(references_txt, (Path(dir_metadata) / "references.txt"))
    writeTxt(references_json, (Path(dir_metadata) / "meta.json"))


def isDocumentOpen(pathstr: str) -> tuple:

    doc_name = Path(pathstr).name
    doc_list = [(i, WD.Documents(i).Name) for i in range(1, WD.Documents.Count+1)]
    for tup in doc_list:
        if tup[1].lower() == doc_name.lower():
            return True, tup[0]
    return False, 0


def run() -> None:
    # commandline interface run command
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

    # add links always on
    processFile(DOC, addHyperlinks=True)
    writeMetaData()

# initialise constants
initialiseDicts()

if __name__ == '__main__':
    run()