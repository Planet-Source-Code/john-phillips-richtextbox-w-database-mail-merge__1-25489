VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Database Mail Merge"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   9360
      TabIndex        =   12
      Top             =   375
      Width           =   9390
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Text            =   "12"
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Text            =   "Arial"
         Top             =   0
         Width           =   1935
      End
      Begin VB.Image Image3 
         Height          =   350
         Left            =   4080
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   350
         Left            =   3600
         Picture         =   "Form1.frx":0102
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   350
         Left            =   3120
         Picture         =   "Form1.frx":0204
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   9360
      TabIndex        =   5
      Top             =   0
      Width           =   9390
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7440
         TabIndex        =   11
         Top             =   0
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4440
         TabIndex        =   9
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Insert Stamp:"
         Height          =   255
         Left            =   6360
         TabIndex        =   10
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Insert data Field:"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Select Table:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9375
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   6360
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   375
         Left            =   8400
         TabIndex        =   2
         Top             =   6360
         Width           =   735
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   5775
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   10186
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0306
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Record 0 of 0"
         Height          =   255
         Left            =   6600
         TabIndex        =   4
         Top             =   6480
         Width           =   1695
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuundo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnucut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnudash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselectall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnumm 
      Caption         =   "&Mail Merge"
      Begin VB.Menu mnuopendata 
         Caption         =   "Open &DataSource"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnumergedata 
         Caption         =   "&Merge Data"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnudash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinsert 
         Caption         =   "&Insert Merge Field"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucontents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnudash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mail merge using a access database and a richtextbox
' by John Phillips, MCP
'
' keep in mind that this can be much much better with more work
' this is just to show how a mail merge using a database on the fly
' and tables can be done.
' it may not look pretty but it does what is needed
' I needed to create an application like this one for
' the place I work at, I just re-did the interface and a few of the functions
' to be uploaded. I am sorry but I cannot upload the completed work
' with the finished interface.
'
' I think I included everything you will need for this example.
' if I forgot something let me know
' if you have any question - email me at vbjack@nyc.rr.com
' oh yeah and have fun codeing

Dim sFilename As String
Dim td As TableDef ' set up base table object in the database
Dim pr As Property ' setup the table properties object
Dim fl As Field ' setup the database field object
Dim multVal As Integer
Dim sFields() As String
Dim nCnt As Integer  ' total record count in database
Dim sTables() As String ' setup the table array

Private Sub Combo3_Click()
If Combo3.Text = "" Then
Combo1.Clear
Exit Sub
End If

Dim recCnt As Integer
Dim x As Integer ' this can be setup in the declerations
' section of the form to eliminate repeative coding
' but for now we will just do it this way

' code to fill the 1st combobox with fields from the selected
' table
' select all the data from the selected table
If RichTextBox1.Text <> "" Then
Dim nResp As Integer
nResp = MsgBox("Selecting a new table will erase your work!" & vbCrLf & "You may only select 1 table to merge at a time" & vbCrLf & "If you want to merge fields from multible table, setup a Query in Access!" & vbCrLf & "Do you want to continue?", vbYesNo + vbQuestion, "Erase Work")
If nResp = 6 Then ' yes code
' erase work and keep going
RichTextBox1.Text = ""
ElseIf nResp = 7 Then ' no code
RichTextBox1.SetFocus
Exit Sub
End If
End If

Data1.RecordSource = "SELECT * FROM " & Combo3.Text
Data1.Refresh
Data1.UpdateControls

Combo1.Clear ' clear the combobox
If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
nCnt = 0
Label1.Caption = "Record 0 of 0"
Else
Data1.Recordset.MoveLast
nCnt = Data1.Recordset.RecordCount
Data1.Recordset.MoveFirst
Label1.Caption = "record " & Data1.Recordset.AbsolutePosition + 1 & " of " & nCnt
End If

Set td = Data1.Database.TableDefs(Combo3.Text) ' setup the table object for the selected table

  For Each fl In td.Fields ' now run through the list of fields and add them to the combobox
    Combo1.AddItem "«" & fl.Name & "»"
  Next
recCnt = Combo1.ListCount

ReDim sFields(recCnt) ' redim the field string array

x = 0
  For Each fl In td.Fields ' now run through the list of fields and add them to the array
    sFields(x) = fl.Name
    x = x + 1
  Next

End Sub

Private Sub Combo4_Click()
' change the font in the richtextbox
   RichTextBox1.SelFontName = Combo4.Text
End Sub

Private Sub Combo5_Click()
If Combo5.Text = "" Then Exit Sub
' change the selected font size
    RichTextBox1.SelFontSize = Combo5.Text
End Sub

Private Sub Command1_Click()
If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then Exit Sub

Data1.Recordset.MoveNext
If Data1.Recordset.EOF = True Then
Data1.Recordset.MovePrevious
End If
fillFields
Label1.Caption = "Record " & Data1.Recordset.AbsolutePosition + 1 & " of " & nCnt

End Sub

Private Sub Command2_Click()
If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then Exit Sub

Data1.Recordset.MovePrevious
If Data1.Recordset.BOF = True Then
Data1.Recordset.MoveNext
End If
fillFields
Label1.Caption = "Record " & Data1.Recordset.AbsolutePosition + 1 & " of " & nCnt

End Sub

Private Sub Form_Load()
Dim i As Integer

With Combo2
.AddItem ""
.AddItem "Time Stamp"
.AddItem "Date Stamp"
End With

' set the combo2 box to the first item in the list
Combo2.ListIndex = 0

' setup the val we need to multiply the slider value to
' to indent the text in the rich textbox control
' first we divide by 10 (the amount of ticks on the slider)
' by the width - then we need to minus 10 from the final value to get an almost exact match
' play with the amount of ticks on the slider control and check out the reults
' also if you change from twips to pixels on the form you must modify these
' calculations
multVal = (RichTextBox1.Width / 10) - 10

' load the fonts into the combobox
' you can change all the combobox's to easyier names to remember
' names if you like, I just find it a waste of time with small projects like this
' thats just my style of programming


With Combo4
   For i = 0 To Screen.FontCount - 1 ' .Count - 1
   .AddItem Screen.Fonts(i)
   Next i
End With

RichTextBox1.SelFontName = Combo4.Text

With Combo5
 For i = 8 To 72 Step 2
    .AddItem i
 Next i
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set td = Nothing
Set pr = Nothing
Set fl = Nothing
End
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1.BorderStyle = 0 Then
' make text bold
Image1.BorderStyle = 1
RichTextBox1.SelBold = True
ElseIf Image1.BorderStyle = 1 Then
' make text normal
Image1.BorderStyle = 0
RichTextBox1.SelBold = False
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Image1.BorderStyle = 0
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image2.BorderStyle = 0 Then
' make text italic
Image2.BorderStyle = 1
RichTextBox1.SelItalic = True
ElseIf Image2.BorderStyle = 1 Then
' make text normal
Image2.BorderStyle = 0
RichTextBox1.SelItalic = False
End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image3.BorderStyle = 0 Then
' make text bold
Image3.BorderStyle = 1
RichTextBox1.SelUnderline = True
ElseIf Image3.BorderStyle = 1 Then
' make text normal
Image3.BorderStyle = 0
RichTextBox1.SelUnderline = False
End If
End Sub

Private Sub mnuabout_Click()
MsgBox "Thank you for downloading this application, if you like it vote for it, if you have something to say about it (good or bad) email me at vbjack@nyc.rr.com"

End Sub

Private Sub mnucontents_Click()
MsgBox "No help files" & vbclrf & vbCrLf & "Click on mail merge then click on open data source - then select an access database" & vbCrLf & "then select the table you want then select and insert the fields that you want", vbOKOnly + vbInformation, "No help files"
MsgBox "Thank you for downloading this application, if you like it vote for it, if you have something to say about it (good or bad) email me at vbjack@nyc.rr.com"

End Sub

Private Sub mnucopy_Click()

   RichTextBox1.SetFocus
   ' Clear the contents of the Clipboard.
   Clipboard.Clear
   ' Copy selected text to Clipboard.
   Clipboard.SetText Screen.ActiveControl.SelText

End Sub

Private Sub mnucut_Click()

   If RichTextBox1.SelText <> "" Then
   RichTextBox1.SetFocus
   RichTextBox1.SelText = ""
   End If

End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuinsert_Click()
If Combo1.Text = "" Then Exit Sub ' make sure something is selected

' it is this simple to insert test into the richtextbox
RichTextBox1.SelBold = True
RichTextBox1.SelText = Combo1.Text
RichTextBox1.SelBold = False
RichTextBox1.SetFocus

End Sub

Private Sub mnumergedata_Click()

   RichTextBox1.SaveFile App.Path & "\tempmerge.rtf"
   
   fillFields
   
End Sub

Private Sub mnuopen_Click()
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  On Error GoTo errHandler
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
  "(*.txt)|*.txt|Rich Text Files (*.rtf)|*.rtf"
  ' Specify default filter
  CommonDialog1.FilterIndex = 2
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  sFilename = CommonDialog1.FileName
  'Exit Sub

  RichTextBox1.LoadFile sFilename
errHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub mnuopendata_Click()
Dim x As Integer
Dim nResp As Integer

If RichTextBox1.Text <> "" Then
nResp = MsgBox("Selecting a new datasource will erase your current work!" & vbCrLf & "Are you sure you want to continue?", vbYesNo + vbQuestion, "Erase Work")
If nResp = 6 Then ' yes code
' do erase work and keep going the user selected to earase the work
RichTextBox1.Text = ""
ElseIf nResp = 7 Then ' no code
'exit this function and return focus to the richtextbox
RichTextBox1.SetFocus
Exit Sub
End If
End If
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  On Error GoTo errHandler
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "Access Database (*.mdb)|*.mdb"
  ' Specify default filter
  CommonDialog1.FilterIndex = 2
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  sFilename = CommonDialog1.FileName
  'Exit Sub
  
  Screen.MousePointer = vbHourglass
  
  Data1.DatabaseName = sFilename
  
  Data1.Refresh   ' Open the Database.
  ' Read and print the name of each table in the database.
  ' to the combobox
  
  Combo3.AddItem ""
  x = 0

  
  For Each td In Data1.Database.TableDefs
    Combo3.AddItem td.Name
    x = x + 1
  Next
  
    ReDim sTables(x) ' redim the tables array
    ReDim sUsedTables(x) ' setup the usedtables array and leave empty for now
    
    x = 0
    
  For Each td In Data1.Database.TableDefs
    sTables(x) = td.Name ' add the tables to the table array
    x = x + 1
  Next
  
  RichTextBox1.Text = ""
  
  Screen.MousePointer = vbNormal
errHandler:
  'User pressed the Cancel button
  Screen.MousePointer = vbNormal
  Exit Sub
End Sub

Private Sub mnupaste_Click()
   
   RichTextBox1.SetFocus
   ' Place text from Clipboard into active control.
   Screen.ActiveControl.SelText = Clipboard.GetText()

End Sub

Private Sub mnusave_Click()
  CommonDialog1.CancelError = True
  On Error GoTo errHandler
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
  "(*.txt)|*.txt|Rich Text Files (*.rtf)|*.rtf"
  ' Specify default filter
  CommonDialog1.FilterIndex = 2
  
If sFilename <> "" Then
   RichTextBox1.SaveFile sFilename
Else
   Dim strNewFile As String
   CommonDialog1.ShowSave
   strNewFile = CommonDialog1.FileName
   RichTextBox1.SaveFile strNewFile
   sFilename = strNewFile
End If
Exit Sub
errHandler:
Exit Sub
' user pressed cancel
End Sub

Private Sub mnusaveas_Click()
   On Error GoTo errHandler
   Dim strNewFile As String
   
   CommonDialog1.CancelError = True
   ' Set flags
   CommonDialog1.Flags = cdlOFNHideReadOnly
   ' Set filters
   CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
   "(*.txt)|*.txt|Rich Text Files (*.rtf)|*.rtf"
   ' Specify default filter
   CommonDialog1.FilterIndex = 2
   CommonDialog1.ShowSave
   strNewFile = CommonDialog1.FileName
   RichTextBox1.SaveFile strNewFile
   sFilename = strNewFile
Exit Sub
errHandler:
Exit Sub
' user pressed cancel

End Sub

Private Sub mnuselectall_Click()

      RichTextBox1.SetFocus ' set focus back to the richtextbox
      RichTextBox1.SelStart = 0  ' set selection start and
      RichTextBox1.SelLength = Len(RichTextBox1.Text) ' selection lenth in this case all

End Sub

Private Sub mnuundo_Click()
MsgBox "This function was not implimented, This example is for the use of merging data not undoing typing. PSC has plenty of examples of how to do this!"
End Sub

Private Sub Slider1_Click()
RichTextBox1.SelIndent = Slider1.Value * multVal
'rtfData.SelIndent = 0.5

End Sub


Private Function fillFields()
On Error GoTo fieldErr
Dim fldCnt As Integer ' setup integer to hold the field count
Dim x As Integer ' setup counter

fldCnt = Combo1.ListCount ' get and set the list count (fields)

If fldCnt = 0 Then Exit Function


RichTextBox1.LoadFile App.Path & "\tempmerge.rtf"

For x = 0 To fldCnt - 1
' here is where the merge takes place
' not to complicated just follow the steps 1 by 1

If Data1.Recordset.Fields(sFields(x)) <> "" Then
    Do While FindMerge("«" & sFields(x) & "»", Data1.Recordset.Fields(sFields(x))) = True

    Loop
Else
    Do While FindMerge("«" & sFields(x) & "»", "") = True

    Loop
End If

Next x

Exit Function
fieldErr:
If Err.Number = 13 Then Resume Next
MsgBox Err.Number & vbCrLf & Err.Description

' most likely the field does not support the value such as a image
' field in a database

'  textfound = RTF.Find(Findme, RTF.SelStart + Len(Findme))
'    If textfound <> -1 Then 'Found it !
'        foundit = True 'let the form know we succeeded
'        RTF.SetFocus 'show it to the user
'        Exit Sub 'lets bail out now
 '   Else
 '       foundit = False 'let the form know we cant find it
 '       Exit Sub 'lets bail out now
 '   End If


End Function


Private Function FindMerge(sFind As String, sReplace As String) As Boolean

RichTextBox1.Find (sFind)
    textfound = RichTextBox1.Find(sFind)
    If textfound <> -1 Then
        FindMerge = True
        'RichTextBox1.SetFocus
        RichTextBox1.SelText = sReplace
        Exit Function
    Else
        FindMerge = False
        Exit Function
    End If

End Function
