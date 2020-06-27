
I have a [`DOCX` file](https://github.com/retsyo/read_word_header_win32com), which has the following info

|           | pages of section | header                                      |
| --------- | ---------------- | ------------------------------------------- |
| section 1 | 1                | None                                        |
| section 2 | 2                | Abstract. current page ??, total IV pages   |
| section 3 | 5                | paper body, current page ??, total 35 pages |

Then if the line `oSec.Headers(1).Range.Fields.Update` is used in the following `VBA` code, the corrected header text will be shown
```vb
Function myTrim(s)
    a = Replace(s, vbLf, "")
    myTrim = Trim(a)
End Function

Sub displayHeader()
    idx = 1
    For Each oSec In ActiveDocument.Sections
		oSec.Headers(1).Range.Fields.Update 'this line must be called
        MsgBox "sec " & idx & " " & myTrim(oSec.Headers(1).Range.Text)
        idx = idx + 1
    Next
End Sub

```

Then I coined the `Python` version, as we all know it looks like the original `VBA` one
```python
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
```
However the `Python` code does not give the same corrected header text:

|                                     | Dispatch('Word.Application')                                 | gencache.EnsureDispatch('Word.Application')                  |
| ----------------------------------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| use `Range.Fields.Update()`         | sec 1 <br/>sec 2 Abstract. current page I, total I pages<br/>sec 3 paper body, current page 1, total 1 pages | sec 1 <br/>sec 2 Abstract. current page I, total I pages<br/>sec 3 paper body, current page 1, total 1 pages |
| do not use  `Range.Fields.Update()` | sec 1 <br/>sec 2 Abstract. current page III, total IV pages<br/>sec 3 paper body, current page 1, total 35 pages | sec 1 <br/>sec 2 Abstract. current page I, total I pages<br/>sec 3 paper body, current page 1, total 35 pages |

So, what is the problem, and how to fix it? Thank you in advance.

