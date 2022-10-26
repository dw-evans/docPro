

import tkinter as tk
from tkinter import filedialog

import utils
import subprocess
from pathlib import Path


LAST_MODIFIED = utils.VersionInfo.last_edited
MODIFIED_BY = '' #utils.VersionInfo.author
VERSION = utils.VersionInfo.version_number
WINDOW_NAME = "documentPro.exe ( •_•)>⌐■-■"


def shortpath(path_string):

    # hack to tidy up Q drive link
    path_string = path_string.replace("\\\\NZ-FS1\\qualityshare\\", 'Q:\\')

    path_arr = path_string.split("\\")
    shortpath_arr = path_arr
    if len(path_arr) > 4:
        shortpath_arr = [path_arr[i] for i in (0,1,0,-2,-1)]
        shortpath_arr[2] = '...'
    shortpath = '\\'.join(shortpath_arr)
    return shortpath


def get_fopen():
    # returns it with forward flashes which need to be replaced later
    f_dir = tk.filedialog.askopenfilename()
    return f_dir

def setLabelText(label_text, text):
    label_text.set(text)
    return

# filetarget = None
# fileoutput = None

class Window:
    def __init__(self, root):

        # target/output file paths (as strings)
        # wd.Documents.Open() doesn't play nice with Path.__repr__
        # hence you'll see Path.resolve.__str__()'s
        self.file_target =  ''  # using '' instead of None for state checks
        self.file_output =  ''

        # Target static label
        self.target_static_text = tk.StringVar()
        self.target_static_text.set("Target:")
        self.target_static_lbl = tk.Label(root, textvariable=self.target_static_text)

        # output static label
        self.output_static_text = tk.StringVar()
        self.output_static_text.set("")
        self.output_static_lbl = tk.Label(root, textvariable=self.output_static_text)

        # target dynamic label (file_target)
        self.target_dynamic_text = tk.StringVar()
        self.target_dynamic_text.set('')
        self.target_dynamic_lbl = tk.Label(root, textvariable=self.target_dynamic_text)

        # output dynamic label (file_output)
        self.output_dynamic_text = tk.StringVar()
        self.output_dynamic_text.set('')
        self.output_dynamic_lbl = tk.Label(root, textvariable=self.output_dynamic_text)

        # static version info footer
        self.version_info_text = tk.StringVar()
        # self.version_info_text.set("(⌐■_■) Version {}. Last updated {} by {}".format(VERSION, LAST_MODIFIED, MODIFIED_BY))
        self.version_info_text.set("(⌐■_■) Version {}. {}".format(VERSION, LAST_MODIFIED))

        self.version_info_lbl = tk.Label(root, textvariable=self.version_info_text)

        # import file button
        self.import_btn = tk.Button(root, text='Import File', command=self.getFileFromDialogue)
        # set file_target using session active document button
        self.activeDocument_btn = tk.Button(root, text='Get Active Document', command=self.getFileFromActive)
        # run (process file) button
        self.run_btn = tk.Button(root, text='Run', command=self.run, width=15)

        # add hyperlinks checkbox
        self.add_hyperlinks = True  # Who wouldn't want this on
        self.add_hyperlinks_val = tk.BooleanVar(value=self.add_hyperlinks)
        self.add_hyperlinks_btn = tk.Checkbutton(root, text="Add Hyperlinks", variable=self.add_hyperlinks_val, command=self.toggleAddHyperlinks) # not self.addHyperlinks)

        # create copy (safe mode) checkbox
        self.create_copy = False
        self.create_copy_val = tk.BooleanVar(value=self.create_copy)
        self.create_copy_btn = tk.Checkbutton(root, text="SAFE MODE", variable=self.create_copy_val, command=self.toggleCreateCopy) # not self.addHyperlinks)

        # write metadata checkbox
        self.write_metadata = True  # maybe remove the choice lol
        self.write_metadata_val = tk.BooleanVar(value=self.write_metadata)
        self.write_metadata_btn = tk.Checkbutton(root, text="Write Metadata", variable=self.write_metadata_val, command=self.toggleWriteMeta) # not self.addHyperlinks)

        # Opens explorer at the target file location, quite handy imo
        self.open_explorer_btn = tk.Button(root, text='Open File Explorer at Target', command=self.openExplorerAtTarget)
        self.open_target_word_btn = tk.Button(root, text='Open Target In Word', command=self.openTargetWord)
        
        self.updateLabels()

        # object placement
        self.activeDocument_btn.place(x=10, y=10)
        self.import_btn.place(x=140, y=10)

        self.add_hyperlinks_btn.place(x=10, y=40)
        self.create_copy_btn.place(x=10, y=60)
        self.write_metadata_btn.place(x=10,y=80)

        self.target_static_lbl.place(x=10, y=110)
        self.target_dynamic_lbl.place(x=60, y=110)

        self.output_static_lbl.place(x=10, y=130)
        self.output_dynamic_lbl.place(x=60, y=130)

        self.run_btn.place(x=10, y=160)

        self.version_info_lbl.place(x=10, y=220)



    def openExplorerAtTarget(self):
        subprocess.Popen('explorer /select, "{}"'.format(self.file_target))

    def update(self):
        # update labels and also add go to explorer button once file selected
        if self.file_target != '':
            self.open_target_word_btn.place(x=10, y=190)
            self.open_explorer_btn.place(x=140, y=190)
            
            if self.create_copy:
                pth = Path(self.file_target)
                self.file_output = (pth.parent / (pth.stem+"_linked"+pth.suffix)).resolve().__str__()


        self.updateLabels()


    def updateLabels(self):
        # updates non static labels
        if self.file_target == '':
            self.target_dynamic_text.set('Select a file')
        else:
            self.target_dynamic_text.set('{}'.format(shortpath(self.file_target)))
            print(shortpath(self.file_target))
        if self.create_copy:
            self.output_static_text.set('Output:')
            if self.file_output != "":
                self.output_dynamic_text.set('{}'.format(shortpath(self.file_output)))
                print(shortpath(self.file_output))
            else:
                self.output_dynamic_text.set('')

        else:
            self.output_static_text.set('')
            self.output_dynamic_text.set('')

    # toggler functions for each checkbox, refreshes window
    def toggleAddHyperlinks(self):
        self.add_hyperlinks = self.add_hyperlinks_val.get()
        self.update()

    def toggleCreateCopy(self):
        self.create_copy = self.create_copy_val.get()
        self.update()

    def toggleWriteMeta(self):
        self.write_metadata = self.write_metadata_val.get()
        self.update()


    def openTargetWord(self):
        # open target file in word
        # would be good to make this bring it to front

        # a = Path(self.file_target).name
        # search = a + ' - Word'
        # print(search)
        # hwnd = win32gui.FindWindow('', search)
        # print(hwnd)

        try:
            utils.WD.Documents.Open(self.file_target)
        except:
            wordErrorHandler()
            # self.openTargetWord()


    def getFileFromActive(self):
        # sets the active word doc to the target
        try:
            doc = utils.WD.ActiveDocument
            self.file_target = (Path(doc.Path) / doc.Name).resolve().__str__()
        except:
            wordErrorHandler()

        self.update()


    
    def run(self):
        # processes the file, only if a file target is set.
        if self.file_target != '':
            utils.initialiseWord()
            print(self.file_target)
            print(self.file_output)

            # need to check if the document is open first, returns true and nonzero idx if already open
            is_open, idx = utils.isDocumentOpen(self.file_target)
            if not is_open:
                utils.WD.Documents.Open(self.file_target)
            else:
                utils.WD.Documents(idx).Activate()
            
            # create_copy handling, either change file or open (and activate) new one
            if not self.create_copy:
                utils.processFile(utils.WD.ActiveDocument, addHyperlinks=self.add_hyperlinks)
            else:
                utils.WD.ActiveDocument.SaveAs2(self.file_output)
                utils.processFile(utils.DOC, addHyperlinks=self.add_hyperlinks)
            
            if self.write_metadata:
                utils.writeMetaData()


   


    def getFileFromDialogue(self):
        # sets the target file using a filedialog
        file_target = get_fopen()
        file_output = ''
        if file_target != '':
            file_target_dirs = file_target.split('/')
            
            # add _linked to outputfile filename
            file_output_dirs = file_target_dirs[:]
            a = file_output_dirs
            file_output_dirs[-1] = '.'.join([a[-1].split('.')[-2]+'_linked', a[-1].split('.')[-1]])
            file_output = '/'.join(file_output_dirs)
            
            self.file_target = file_target.replace('/','\\')
            self.file_output = file_output.replace('/','\\')

        global filetarget, fileoutput
        filetarget = self.file_target
        fileoutput = self.file_output

        self.update()

def wordErrorHandler():
    # errors with things disconnecting, reinitialising seems to work
    print("Error interacting with the Word session, starting another...")
    utils.initialiseWord()


def main():
    # main func
    root = tk.Tk()
    win = Window(root)
    root.title(WINDOW_NAME)
    root.geometry("480x250+10+10")
    root.mainloop()

if __name__ == "__main__":
    main()