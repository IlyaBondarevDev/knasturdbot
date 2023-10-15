import win32com.client as wc
from constants import MACROS_CODE, SUB_NAME

class MacroExecuter:
    def Execute(this, path):       
        word = wc.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        macro = doc.VBProject.VBComponents.Add(1)
        macro.CodeModule.AddFromString(MACROS_CODE)
        word.Run(SUB_NAME)
        doc.VBProject.VBComponents.Remove(macro)
        doc.Save()
        doc.Close()
        word.Quit()

            
MacroExecuter().Execute("C:\\AllMain\\Programming\\Telegram\\knasturdbot\\bot\\files\\787314003\\ТСП 2.2.docx")