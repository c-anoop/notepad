VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNotepad 
   Caption         =   " Untitled - Notepad"
   ClientHeight    =   4950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6615
   Icon            =   "frmNotepad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog wcdNote 
      Left            =   1080
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox rtbNote 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8705
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   20000
      TextRTF         =   $"frmNotepad.frx":030A
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu submnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu submnuFileOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu submnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu submnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu submnuFilePageSetup 
         Caption         =   "Page Setup..."
      End
      Begin VB.Menu submnuFilePrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu submnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu submnuEditUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu submnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu submnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu submnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu submnuEditDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu submnuEditFind 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu submnuEditFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu submnuEditreplace 
         Caption         =   "Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu submnuEditGoTo 
         Caption         =   "Go To..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu submnuEditSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu submnuEditTimeDate 
         Caption         =   "Time/Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Begin VB.Menu submnuFormatWordWrap 
         Caption         =   "Word Wrap"
      End
      Begin VB.Menu submnuFormatFont 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu submnuViewStatusBar 
         Caption         =   "Status Bar"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu submnuHelpHelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu submnuHelpAboutNotepad 
         Caption         =   "About Notepad"
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu submnuContextUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu submnuContextCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu submnuContextCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu submnuContextPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu submnuContextDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuDash8 
         Caption         =   "-"
      End
      Begin VB.Menu submnuContextSelectAll 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "frmNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Notepad's Code Starts Here

Option Explicit

Dim boolvalid1 As Boolean
Dim strString1 As String
Dim strString2 As String
Dim boolValid2 As Boolean

Private Sub Form_Load()

 ' This subroutine sets the Enabled properties of menus
   submnuEditUndo.Enabled = False
   submnuEditCut.Enabled = False
   submnuEditCopy.Enabled = False
   submnuEditDelete.Enabled = False
   submnuContextUndo.Enabled = False
   submnuContextCut.Enabled = False
   submnuContextDelete.Enabled = False
   submnuContextSelectAll.Enabled = False
   submnuContextCopy.Enabled = False
   submnuEditSelectAll.Enabled = False
   strString1 = ""
   strString2 = ""
   
End Sub

Private Sub Form_Resize()
   
 ' This subroutine is for setting the size of _
   RichTextBox according to the Form size
   If frmNotepad.Height > 800 Then
      rtbNote.Width = frmNotepad.Width - 100
      rtbNote.Height = frmNotepad.Height - 800
   End If

End Sub

Private Sub rtbNote_Change()
 
 ' Sets the Flag boolValid1 to False when text changes _
   in the RitchTextBox
   boolvalid1 = False
   submnuEditUndo.Enabled = True
   submnuContextUndo.Enabled = True
   If rtbNote.Text <> "" Then
      submnuEditSelectAll.Enabled = True
      submnuContextSelectAll.Enabled = True
   Else
      submnuEditSelectAll.Enabled = True
      submnuContextSelectAll.Enabled = True
   End If
   
End Sub



Private Sub rtbNote_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

 ' This Subroutine Displays the Context Menu on Right _
   click of the mouse on RitchTextBox
   If Button = vbRightButton Then
      PopupMenu mnuContext
   End If

End Sub

Private Sub rtbNote_SelChange()
  
 ' This subroutine sets the menu Enabled Properties
 ' This part runs when there is some selected text
   If Len(rtbNote.SelText) > 0 Then
      submnuEditCut.Enabled = True
      submnuEditCopy.Enabled = True
      submnuEditDelete.Enabled = True
      submnuContextCopy.Enabled = True
      submnuContextDelete.Enabled = True
      submnuContextCut.Enabled = True
    Else ' This part runs when there is no selected text
      submnuEditCut.Enabled = False
      submnuEditCopy.Enabled = False
      submnuEditDelete.Enabled = False
      submnuContextCopy.Enabled = False
      submnuContextDelete.Enabled = False
      submnuContextCut.Enabled = False
   End If

End Sub

Private Sub submnuContextCopy_Click()
 
 ' This subroutine runs when Copy of the shortcut _
   menu is clicked
   submnuEditCopy_Click

End Sub

Private Sub submnuContextCut_Click()

 ' This subroutine runs when Cut of the shortcut _
   menu is clicked
   submnuEditCut_Click

End Sub

Private Sub submnuContextDelete_Click()

 ' This subroutine runs when delete of the shortcut _
   menu is clicked
   submnuEditDelete_Click

End Sub

Private Sub submnuContextPaste_Click()

 ' This subroutine runs when Paste of the shortcut _
   menu is clicked
   submnuEditPaste_Click

End Sub

Private Sub submnuContextSelectAll_Click()
 
 ' This subroutine runs when Select All of the shortcut _
   menu is clicked
   submnuEditSelectAll_Click
  
End Sub

Private Sub submnuContextUndo_Click()
 
 ' This subroutine runs when Undo of the shortcut _
   menu is clicked
   submnuEditUndo_Click

End Sub

Private Sub submnuEditCopy_Click()

 ' This subroutine stores the selected text to _
   Windows Clipboard
   Clipboard.SetText (rtbNote.SelText)

End Sub

Private Sub submnuEditCut_Click()

 ' This subroutine stores the selected text to Windows _
   Clipboard and deletes it
   Clipboard.SetText (rtbNote.SelText)
   strString1 = rtbNote.SelText
   rtbNote.SelText = ""
   boolValid2 = True

End Sub

Private Sub submnuEditDelete_Click()

 ' This subroutine deletes the selected text
   rtbNote.SelText = ""

End Sub

Private Sub submnuEditFind_Click()
  
 ' This subroutine finds a given string in the text
  Dim strFindString As String
  Dim lngPosition As Long
   strFindString = InputBox("Find what:", "Find")
   
   If rtbNote.Text = "" Then
      Exit Sub
   End If
      
   lngPosition = rtbNote.Find(strFindString, 0)
   If lngPosition > 0 And Mid(rtbNote.Text, _
                          lngPosition, 1) = " " Then
       rtbNote.SelStart = lngPosition
                            
       rtbNote.SelLength = Len(strFindString)
   Else
       Call MsgBox("Can't find" & " " & strFindString, _
                     vbOKOnly, "Find")
   End If
   
End Sub

Private Sub submnuEditFindNext_Click()

 ' This subroutine runs when Find Next is clicked
 ' It finds the next word after the previously Found Word
   Beep
   MsgBox "No code! Please enter code in subroutine submnuEditFindNext_click()"


End Sub

Private Sub submnuEditGoTo_Click()

   Beep
   MsgBox "No code! Please enter code in subroutine submnuEditGoTo_click()"

End Sub

Private Sub submnuEditPaste_Click()

 ' This subroutine runs when paste is clicked _
   it gets the Text form the Windows Clipboard
   rtbNote.SelText = Clipboard.GetText
 
End Sub

Private Sub submnuEditreplace_Click()

   Beep
   MsgBox "No code! Please enter code in subroutine submnuEditReplace_click()"



End Sub

Private Sub submnuEditSelectAll_Click()

 ' This subroutine selects all text
   rtbNote.SelStart = 0
   rtbNote.SelLength = Len(rtbNote.Text)

End Sub

Private Sub submnuEditTimeDate_Click()

 ' This subroutone tells the time and date
   rtbNote.Text = rtbNote.Text + " " & Str(Time()) & "   " _
                  & Str(Date)
  
End Sub

Private Sub submnuEditUndo_Click()

 ' This subroutine is for Undo
   rtbNote.SelStart = 0
   strString2 = rtbNote.Text
   If boolValid2 = True Then
      rtbNote.Text = rtbNote.Text + strString1
      boolValid2 = False
   Else
      rtbNote.Text = strString1
   End If
   rtbNote.SelLength = Len(rtbNote.Text)
   strString1 = strString2

End Sub

Private Sub submnuFileExit_Click()

 ' This Subroutine runs when exit is clicked _
   and it checks for the unsaved Text
  Dim strAns As String
   If boolvalid1 = False And rtbNote.Text <> "" Then
      strAns = MsgBox("Save Changes?", vbExclamation + _
                     vbYesNo, "Notepad")
      If strAns = vbYes Then
         submnuFileSave_Click
      End If
   End If
 ' Unloads Form
   Unload Me
   Set frmNoteHelp = Nothing
   Set frmNoteHelp = Nothing
   Set frmNotepad = Nothing
   Set frmNotepad = Nothing
   
End Sub

Private Sub submnuFileNew_Click()
 
 ' This Subroutine runs when New is clicked in the menu _
   it checks for unsaved Text
  Dim strAns As String
   If boolvalid1 = False And rtbNote.Text <> "" Then
      strAns = MsgBox("Save Changes?", vbExclamation + _
                     vbYesNoCancel, "Notepad")
      If strAns = vbYes Then
         submnuFileSave_Click
      ElseIf strAns = vbCancel Then
         Exit Sub
      End If
   End If
   rtbNote.Text = ""
 ' Changes the Caption of the Form
   frmNotepad.Caption = " Untitled - Notepad"
   Form_Load

End Sub

Private Sub submnuFileOpen_Click()
 
 ' This Subroutine runs when Open is clicked in the menu _
   it checks for the unsaved Text
  Dim strAns As String
 ' Starts the Error Handler
   On Error GoTo errhandler
   If boolvalid1 = False And rtbNote.Text <> "" Then
      strAns = MsgBox("Save Changes?", vbExclamation + _
                     vbYesNoCancel, "Notepad")
      If strAns = vbYes Then
         submnuFileSave_Click
      ElseIf strAns = vbCancel Then
         Exit Sub
      End If
   End If
 ' Sets the Windows Common Dialog Filter to following _
   Extensions
   wcdNote.Filter = "Text Documents(*.txt)|*.txt" & _
                    "|All Files (*.*)|*.*"
 ' Sets the Common Dialog Flags
   wcdNote.Flags = cdlOFNHideReadOnly Or _
                   cdlOFNNoChangeDir
   wcdNote.ShowOpen
   rtbNote.LoadFile (wcdNote.FileName)
   procTitle
   boolvalid1 = True
 
 ' Error Handler checks for the error=32755 which _
   is generated due to clicking Cancel Button
errhandler:    If Err.Number = 32755 Then
                  Exit Sub
               End If
End Sub

Private Sub submnuFilePageSetup_Click()

 ' Page Setup code goes here
   Beep
   MsgBox "No code! Please enter code in subroutine submnuFilePageSetup_click()"

End Sub

Private Sub submnuFilePrint_Click()

 ' This subroutine runs when print is clicked
 ' Printer code should be written here
   Beep
   MsgBox "No code! Please enter code in subroutine submnuFilePrint_click()"

End Sub

Private Sub submnuFileSave_Click()

 ' This subroutine runs when save button is clicked
 ' Error Handler Starts here
   On Error GoTo errhandler
 ' Sets the Default extension to "txt"
   wcdNote.DefaultExt = "txt"
 ' Checks wether the document is new or old such that _
   it opens Save Dialog Box
   If frmNotepad.Caption = " Untitled - Notepad" Then
      wcdNote.ShowSave
      procTitle
   End If
   rtbNote.SaveFile (wcdNote.FileName)
   boolvalid1 = True
  
 ' Checks for the Error 32755
errhandler: If Err.Number = 32755 Then
               Exit Sub
            End If
 
End Sub

Private Sub submnuFileSaveAs_Click()
 
 ' This subroutine runs when Save As is Clicked on the _
   menu
 ' Starts the Error Handler
   On Error GoTo errhandler
 ' sets the Default Extension to "txt"
   wcdNote.DefaultExt = "txt"
   wcdNote.ShowSave
   procTitle
   rtbNote.SaveFile (wcdNote.FileName)
   boolvalid1 = True
   
 ' Checks for the error 32755
errhandler: If Err.Number = 32755 Then
               Exit Sub
            End If
End Sub

Private Sub procTitle()

 ' This subroutine sets the caption of the Document
  Dim strTitle As String
  Dim bytStrLength As Byte
 ' Checks for the "." in the filename i.e. there _
   exists an extension and if exists it removes _
   last 4 chars including "." and extension and _
   displays in the caption of the Form
   If InStr(1, wcdNote.FileTitle, ".") Then
      bytStrLength = Len(wcdNote.FileTitle)
      strTitle = Left(wcdNote.FileTitle, _
      bytStrLength - 4)
   Else
      strTitle = wcdNote.FileTitle
   End If
   frmNotepad.Caption = " " & strTitle & " - Notepad "

End Sub

Private Sub submnuFormatFont_Click()
   
 ' This subroutine loads the Windows Common Font Dialog
   On Error GoTo errhandler
   wcdNote.Flags = cdlCFApply Or cdlCFBoth Or cdlCFEffects _
                  Or cdlCFLimitSize
 ' sets the Font Dialogs Max and Min Sise properties
   wcdNote.Min = 8
   wcdNote.Max = 28
   wcdNote.ShowFont
 ' Sets the RitchTextBox Font Options as selected in the _
   Font Dialog
   rtbNote.SelFontName = wcdNote.FontName
   rtbNote.SelBold = wcdNote.FontBold
   rtbNote.SelItalic = wcdNote.FontItalic
   rtbNote.SelFontSize = wcdNote.FontSize
   rtbNote.SelStrikeThru = wcdNote.FontStrikethru
   rtbNote.SelUnderline = wcdNote.FontUnderline
   rtbNote.SelColor = wcdNote.Color
   
 ' Checks for error 32755
errhandler: If Err.Number = 32755 Then
               Exit Sub
            End If

End Sub

Private Sub submnuFormatWordWrap_Click()

 ' Word Wrap code goes here
   Beep
   MsgBox "No code! Please enter code in subroutine submnuFormatWordWrap_click()"


End Sub

Private Sub submnuHelpAboutNotepad_Click()

 ' Discription about Notepad goes here
  
 ' Beep
 ' MsgBox "No code! Please enter code in subroutine submnuHelpAboutNotepad_click()"
   Load frmNoteHelp
   frmNoteHelp.Visible = True
 
 '                             Code written by
 
 '                             Anoop Chargotra
 '                             Jammu (J&K)
 '                             Wednesday, Nov 06,2002

End Sub

Private Sub submnuHelpHelpTopics_Click()
  
 ' Help Topics code goes here
   Beep
   MsgBox "No code! Please enter code in subroutine submnuHelpTopics_click()"
  
End Sub

Private Sub submnuViewStatusBar_Click()

 ' Status Bar Code goes here
 ' At present there is no Status Bar
   Beep
   MsgBox "No code! Please enter code in subroutine submnuViewStatusBar_click()"

End Sub
