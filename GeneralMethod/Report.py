import os, shutil
from win32com import client  # pip install pywin32
# from docx import Document

demo_replace_dict = {
    "1": "1.123", 
    "2": "2.234",
    "x": "996"
}

class Report:
    def __init__(self):
        self.replace_words = {}

    def load_replace_kw(self, rep: dict) :
        self.replace_words = rep

    def fill_report(self, in_fname, out_fname):
        shutil.copy(in_fname, out_fname) # need full name
        word = client.gencache.EnsureDispatch("Word.Application")
        word.Visible = 0
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.getcwd() + "/" + out_fname)
        word.Selection.Find.ClearFormatting()
        word.Selection.Find.Replacement.ClearFormatting()
        for rep_key in self.replace_words.keys():
            # print("Replacing #%s# to %s" % (rep_key, self.replace_words[rep_key]))
            word.Selection.Find.Execute( '#'+rep_key+'#' ,False,False,False,False,False,True,client.constants.wdFindContinue,False,self.replace_words[rep_key],client.constants.wdReplaceAll)
        doc.Close()
        word.Quit()


if __name__ == '__main__':
    RW = Report()
    RW.load_replace_kw(demo_replace_dict)
    RW.fill_report("../Report/Demo/demo.docx", "demo_filled.docx")