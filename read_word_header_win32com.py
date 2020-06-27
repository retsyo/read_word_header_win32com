import win32com
from win32com.client import Dispatch, constants

#~ word = win32com.client.Dispatch('Word.Application')

word = win32com.client.gencache.EnsureDispatch('Word.Application')

word.Visible = 1
word.DisplayAlerts = 0

word.Documents.Open('r:/test.docx')

for idx, oSec in enumerate(word.ActiveDocument.Sections):
    #~ oSec.Headers(1).Range.Fields.Update()
    print(f'sec {idx+1}', oSec.Headers(1).Range.Text.strip())

word.Documents.Close(constants.wdDoNotSaveChanges)
word.Quit()
