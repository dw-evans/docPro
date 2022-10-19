

import tkinter as tk
from tkinter import filedialog

import utils
import subprocess
from pathlib import Path

from pywintypes import com_error

LAST_MODIFIED = utils.VersionInfo.last_edited
MODIFIED_BY = utils.VersionInfo.author
VERSION = utils.VersionInfo.version_number
WINDOW_NAME = "documentPro.exe  ( •_•)>⌐■-■  >>  (⌐■_■)"


def shortpath(path_string):
    # returns shortened path string 
    path_arr = path_string.split('\\')
    shortpath_arr = path_arr
    if len(path_arr) > 5:
        shortpath_arr = [path_arr[i] for i in (0,1,0,-3,-2,-1)]
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

        self.file_target =  ''
        self.file_output =  ''

        self.target_static_text = tk.StringVar()
        self.target_static_text.set("Target:")
        self.target_static_lbl = tk.Label(root, textvariable=self.target_static_text)

        self.output_static_text = tk.StringVar()
        self.output_static_text.set("")
        self.output_static_lbl = tk.Label(root, textvariable=self.output_static_text)

        self.target_dynamic_text = tk.StringVar()
        self.target_dynamic_text.set('')

        self.output_dynamic_text = tk.StringVar()
        self.output_dynamic_text.set('')

        self.version_info_text = tk.StringVar()
        self.version_info_text.set("Version {}. Last updated {} by {}".format(VERSION, LAST_MODIFIED, MODIFIED_BY))
        self.version_info_lbl = tk.Label(root, textvariable=self.version_info_text)

        self.import_btn = tk.Button(root, text='Import File', command=self.getFileFromDialogue)
        self.activeDocument_btn = tk.Button(root, text='Get Active Document', command=self.getFileFromActive)

        self.target_dynamic_lbl = tk.Label(root, textvariable=self.target_dynamic_text)
        self.output_dynamic_lbl = tk.Label(root, textvariable=self.output_dynamic_text)

        self.run_btn = tk.Button(root, text='Run', command=self.run, width=15)

        

        self.add_hyperlinks = True
        self.add_hyperlinks_val = tk.BooleanVar(value=self.add_hyperlinks)
        self.add_hyperlinks_btn = tk.Checkbutton(root, text="Add Hyperlinks", variable=self.add_hyperlinks_val, command=self.toggleAddHyperlinks) # not self.addHyperlinks)

        self.create_copy = False
        self.create_copy_val = tk.BooleanVar(value=self.create_copy)
        self.create_copy_btn = tk.Checkbutton(root, text="SAFE MODE", variable=self.create_copy_val, command=self.toggleCreateCopy) # not self.addHyperlinks)

        self.write_metadata = True
        self.write_metadata_val = tk.BooleanVar(value=self.write_metadata)
        self.write_metadata_btn = tk.Checkbutton(root, text="Write Metadata", variable=self.write_metadata_val, command=self.toggleWriteMeta) # not self.addHyperlinks)


        self.open_explorer_btn = tk.Button(root, text='Open File Explorer at Target', command=self.openExplorerAtTarget)
        self.open_target_word_btn = tk.Button(root, text='Open Target In Word', command=self.openTargetWord)
        
        self.updateLabels()

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
        # update dynamic labels
        if self.file_target == '':
            self.target_dynamic_text.set('Select a file')
        else:
            self.target_dynamic_text.set('{}'.format(shortpath(self.file_target)))
        if self.create_copy:
            self.output_static_text.set('Output:')
            if self.file_output != "":
                self.output_dynamic_text.set('{}'.format(shortpath(self.file_output)))
            else:
                self.output_dynamic_text.set('')

        else:
            self.output_static_text.set('')
            self.output_dynamic_text.set('')

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
        try:
            utils.WD.Documents.Open(self.file_target)
            utils.WD.Activate()
        except:
            wordErrorHandler()
            self.openTargetWord()


    def getFileFromActive(self):
        try:
            doc = utils.WD.ActiveDocument
            self.file_target = (Path(doc.Path) / doc.Name).resolve().__str__()
        except:
            wordErrorHandler()


        self.update()


    
    def run(self):
        if self.file_target != '':
            utils.initialiseWord()
            print(self.file_target)
            print(self.file_output)
            if not self.create_copy:
                utils.WD.Documents.Open(self.file_target)
                utils.processFile(utils.WD.ActiveDocument, addHyperlinks=self.add_hyperlinks)
            else:
                utils.WD.Documents.Open(self.file_target, ReadOnly=True)
                utils.WD.ActiveDocument.SaveAs2(self.file_output)
                utils.processFile(utils.WD.ActiveDocument, addHyperlinks=self.add_hyperlinks)
                # utils.WD.Documents.Open(self.file_target)
                # utils.WD.Documents.Open(self.file_output)
            
            if self.write_metadata:
                utils.writeMetaData()


   


    def getFileFromDialogue(self):
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
    print("Error interacting with the Word session, starting another...")
    utils.initialiseWord()


def main():
    root = tk.Tk()
    win = Window(root)
    root.title(WINDOW_NAME)
    root.geometry("480x250+10+10")
    root.mainloop()

if __name__ == "__main__":
    main()