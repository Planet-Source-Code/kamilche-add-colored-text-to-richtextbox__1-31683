<div align="center">

## Add Colored Text to RichTextbox


</div>

### Description

Adds the text to a textbox, checking for length overflow. First, it checks, and if the textbox exceeds 15,000 characters, it strips out all but the last 2000 characters to make room for the new text. It doesn't break in the middle of lines - it only deletes 'whole lines.' Then it adds the text to the end, using the color you specified (if any), and scrolls to the end of the textbox. I use it all the time, got tired of cutting/pasting out of old projects, thought I'd put it here on PSC.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kamilche](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kamilche.md)
**Level**          |Beginner
**User Rating**    |4.9 (64 globes from 13 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kamilche-add-colored-text-to-richtextbox__1-31683/archive/master.zip)





### Source Code

```
Public Sub Display(ByVal s As String, Optional Color As Long = vbGreen)
  'Add text to the text output window.
  With frmMain.RichTextBox1
    'Clear all but the last 2000 characters if it's too large
    '(don't cut it off in the middle of a line tho).
    If Len(.Text) + Len(s) > 15000 Then
      .SelStart = 0
      .SelLength = InStrRev(.Text, vbCrLf, Len(.Text) - 2000, vbTextCompare) + 1
      .SelText = ""
    End If
    .SelStart = Len(.Text)
    .SelColor = Color
    .SelText = s & vbCrLf
    .SelStart = Len(.Text)
  End With
End Sub
```

