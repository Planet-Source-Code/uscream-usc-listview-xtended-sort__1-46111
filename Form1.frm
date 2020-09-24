VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "uScream (uscream@vip.hr)"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd1k 
      Caption         =   "Add 1000"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add Item"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show/Hide Extra Column"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin MSComctlLib.ImageList imglstListImages 
      Left            =   0
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   9
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":005C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "imglstListImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Desc 
      BackStyle       =   0  'Transparent
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   8160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LVXS As New uSc_ListView_Xtended_Sort
    'Create one uSc_ListView_Xtended_Sort
    'object for each VistView Control and
    'Set it um with SetUp Method
    '(see Form_Load)

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVXS.Sort True, ColumnHeader.Index, Val(ColumnHeader.Tag)
    'if you can't use ColumnHeader.Tag for
    'sort type, you can use integer array instead
End Sub

Private Sub cmdAddItem_Click()
    LVXS.Sort False
    LW_AddItem
    LVXS.Sort True
    'Turm off sorting before adding an item
End Sub

Private Sub cmdAdd1k_Click()
    LVXS.Sort False
    For i = 1 To 1000
        LW_AddItem
    Next
    LVXS.Sort True
End Sub


Private Sub Form_Load()
LW_CreateColumns

    ' You need one extra column
ListView1.ColumnHeaders.Add , , "---Xtra---"
ListView1.ColumnHeaders(8).Width = 2000
 
LVXS.SetUp ListView1

Call Command1_Click

End Sub


'############################################
'#      Crap Below is not important         #
'#    (Except the ColumnHeaders Tags!)      #
'############################################


Private Sub Form_Unload(Cancel As Integer)
ListView1.ListItems.Clear
Unload Me
End Sub

Private Sub Command1_Click()
ListView1.ColumnHeaders(8).Width = IIf(ListView1.ColumnHeaders(8).Width = 0, 2000, 0)
End Sub


Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If x < ListView1.ColumnHeaders(1).Width Then
    Desc.Caption = "Title sorting - it disregard 'The' and 'A' prefixes" _
    & vbNewLine & "SortType = 1"
ElseIf x < ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(1).Width Then
    Desc.Caption = "Integer sorting - Faster then Long Integer sorting" _
    & vbNewLine & "SortType = 2"
ElseIf x < ListView1.ColumnHeaders(3).Width + ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(1).Width Then
    Desc.Caption = "Long Integer sorting" _
    & vbNewLine & "SortType = 3"
ElseIf x < ListView1.ColumnHeaders(4).Width + ListView1.ColumnHeaders(3).Width + ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(1).Width Then
    Desc.Caption = "Decimal Number sorting - formated ####.####" _
    & vbNewLine & "SortType = 4"
ElseIf x < ListView1.ColumnHeaders(5).Width + ListView1.ColumnHeaders(4).Width + ListView1.ColumnHeaders(3).Width + ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(1).Width Then
    Desc.Caption = "Date sorting - Formatted dd. mm. yyyy" _
    & vbNewLine & "SortType = 5"
ElseIf x < ListView1.ColumnHeaders(6).Width + ListView1.ColumnHeaders(5).Width + ListView1.ColumnHeaders(4).Width + ListView1.ColumnHeaders(3).Width + ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(1).Width Then
    Desc.Caption = "Time sorting - Formatted hh:mm" _
    & vbNewLine & "SortType = 6"
ElseIf x < ListView1.ColumnHeaders(7).Width + ListView1.ColumnHeaders(6).Width + ListView1.ColumnHeaders(5).Width + ListView1.ColumnHeaders(4).Width + ListView1.ColumnHeaders(3).Width + ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(1).Width Then
    Desc.Caption = "Normal sorting - if you wanna fast alphabetic sorting in some columns" _
    & vbNewLine & "SortType = 0"
Else
    Desc.Caption = ""
End If
End Sub

Private Sub LW_CreateColumns()
Dim TempString As String
Dim intChr As Integer
Dim intPrefix As Integer

ListView1.ColumnHeaders.Add , , "Title"
ListView1.ColumnHeaders(1).Tag = 1 'Title
ListView1.ColumnHeaders(1).Width = 1600

ListView1.ColumnHeaders.Add , , "Integer"
ListView1.ColumnHeaders(2).Tag = 2 'Integer
ListView1.ColumnHeaders(2).Width = 1000

ListView1.ColumnHeaders.Add , , "Long"
ListView1.ColumnHeaders(3).Tag = 3 'Long
ListView1.ColumnHeaders(3).Width = 1300

ListView1.ColumnHeaders.Add , , "Decimal (Single/Double)"
ListView1.ColumnHeaders(4).Tag = 4 'Long
ListView1.ColumnHeaders(4).Width = 1300

ListView1.ColumnHeaders.Add , , "Date"
ListView1.ColumnHeaders(5).Tag = 5 'Date
ListView1.ColumnHeaders(5).Width = 1300

ListView1.ColumnHeaders.Add , , "Time"
ListView1.ColumnHeaders(6).Tag = 6 'Time
ListView1.ColumnHeaders(6).Width = 600

ListView1.ColumnHeaders.Add , , "Normal Sort (Faster)"
ListView1.ColumnHeaders(7).Tag = 0 'Normal
ListView1.ColumnHeaders(7).Width = 1600

End Sub


Private Sub LW_AddItem()
    TempString = ""
    For j = 1 To Int(Rnd * 10) + 1
        TempString = TempString & Chr((Int(Rnd * 10) * 2) + 65)
    Next
    TempString = LCase(TempString)
    intPrefix = Int(Rnd * 5 + 1)
    If intPrefix = 1 Then TempString = "The " & TempString
    If intPrefix = 2 Then TempString = "A " & TempString
    
    
    ListView1.ListItems.Add , , TempString

'INTEGER
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = Int((Rnd * 60) * 1000) - 30000
'LONG
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = Int((Rnd * 4) * 1000000000) - 2000000000
'DECIMAL
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = (Rnd * 4) * 100 - 200
'DATE
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = CInt((Rnd * 27) + 1) & ". " & CInt((Rnd * 11) + 1) & ". " & CInt((Rnd * 1100) + 1000)
'TIME
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = CInt((Rnd * 22) + 1) & ":" & CInt((Rnd * 58) + 1)

'NORMAL SORT
    TempString = ""
    For j = 1 To Int(Rnd * 10) + 1
        TempString = TempString & Chr((Int(Rnd * 10) * 2) + 65)
    Next
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = TempString

End Sub


Private Sub cmdClear_Click()
ListView1.ListItems.Clear
End Sub
