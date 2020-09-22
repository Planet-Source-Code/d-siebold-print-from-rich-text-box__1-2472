<div align="center">

## Print from rich text box


</div>

### Description

This code combines the excellent submissions of PrintCode by Ken Chia and the printing from a rich text control by VBPro to display the Windows printing common dialog and then send the rich text box contents to the selected printer with formatting and margins.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[D\. Siebold](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/d-siebold.md)
**Level**          |Unknown
**User Rating**    |5.0 (55 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/d-siebold-print-from-rich-text-box__1-2472/archive/master.zip)





### Source Code

```
'This is where the printing is called - assumes a form or UserControl with Windows common dialog control called dlgPrint, a rich text box called rtbText and a command button called cmdPrint
Private Sub cmdPrint_Click()
  dlgPrint.Flags = cdlPDReturnDC + cdlPDNoPageNums
  If rtbText.SelLength = 0 Then
    dlgPrint.Flags = dlgPrint.Flags + cdlPDAllPages
  Else
    dlgPrint.Flags = dlgPrint.Flags + cdlPDSelection
  End If
  dlgPrint.ShowPrinter
  PrintRTF rtbText, 1440, 1440, 1440, 1440 ' 1440 Twips = 1 Inch
End Sub
'Printing constants - these should go in form or UserControl Declarations
Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type CharRange
  cpMin As Long    ' First character of range (0 For start of doc)
  cpMax As Long    ' Last character of range (-1 For End of doc)
End Type
Private Type FormatRange
  hdc As Long     ' Actual DC to draw on
  hdcTarget As Long  ' Target DC For determining text formatting
  rc As Rect     ' Region of the DC to draw to (in twips)
  rcPage As Rect   ' Region of the entire DC (page size) (in twips)
  chrg As CharRange  ' Range of text to draw (see above declaration)
End Type
Private Declare Function GetDeviceCaps Lib "gdi32" ( _
  ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
  lp As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
  (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
  ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
'Routine that does the printing
Private Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, _
          RightMarginWidth, BottomMarginHeight)
  On Error GoTo ErrorHandler
  Dim LeftOffset As Long, TopOffset As Long
  Dim LeftMargin As Long, TopMargin As Long
  Dim RightMargin As Long, BottomMargin As Long
  Dim fr As FormatRange
  Dim rcDrawTo As Rect
  Dim rcPage As Rect
  Dim TextLength As Long
  Dim NextCharPosition As Long
  Dim R As Long
  ' Start a print job to get a valid Printer.hDC
  Printer.Print Space(1)
  Printer.ScaleMode = vbTwips
  ' Get the offsett to the printable area on the page in twips
  LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
  PHYSICALOFFSETX), vbPixels, vbTwips)
  TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
  PHYSICALOFFSETY), vbPixels, vbTwips)
  ' Calculate the Left, Top, Right, and Bottom margins
  LeftMargin = LeftMarginWidth - LeftOffset
  TopMargin = TopMarginHeight - TopOffset
  RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
  BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset
  ' Set printable area rect
  rcPage.Left = 0
  rcPage.Top = 0
  rcPage.Right = Printer.ScaleWidth
  rcPage.Bottom = Printer.ScaleHeight
  ' Set rect in which to print (relative to printable area)
  rcDrawTo.Left = LeftMargin
  rcDrawTo.Top = TopMargin
  rcDrawTo.Right = RightMargin
  rcDrawTo.Bottom = BottomMargin
  ' Set up the print instructions
  fr.hdc = Printer.hdc ' Use the same DC For measuring and rendering
  fr.hdcTarget = Printer.hdc ' Point at printer hDC
  fr.rc = rcDrawTo ' Indicate the area On page to draw to
  fr.rcPage = rcPage ' Indicate entire size of page
  fr.chrg.cpMin = 0 ' Indicate start of text through
  fr.chrg.cpMax = -1 ' End of the text
  ' Get length of text in RTF
  TextLength = Len(RTF.Text)
  ' Loop printing each page until done
  Do
    ' Print the page by sending EM_FORMATRANGE message
    NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
    If NextCharPosition >= TextLength Then Exit Do 'If done then exit
    fr.chrg.cpMin = NextCharPosition ' Starting position For next page
    Printer.NewPage ' Move On to Next page
    Printer.Print Space(1) ' Re-initialize hDC
    fr.hdc = Printer.hdc
    fr.hdcTarget = Printer.hdc
  Loop
  ' Commit the print job
  Printer.EndDoc
  ' Allow the RTF to free up memory
  R = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
ErrorHandler:
End Sub
```

