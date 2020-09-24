VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "uSc_ListView_Xtended_Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************
'*  Copyright (C) uScream 2003 - All Rights Reserved    *
'*                                                      *
'*  Contact: uscream@vip.hr                             *
'*                                                      *
'*  CHANGE HISTORY:                                     *
'*      10.06.2003. - v 1.0                             *
'*      15.07.2003. - v 1.1                             *
'*          - Sort icon on right side                   *
'*              - function from ListView Each Row Tooltip (http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46801&lngWId=1)
'*                                                      *
'*  TESTED WITH: mscomctl.ocx ver. 6.00.8862            *
'*                                                      *
'********************************************************
Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Const HDS_HOTTRACK = &H4

Private Const HDI_BITMAP = &H10
Private Const HDI_IMAGE = &H20
Private Const HDI_ORDER = &H80
Private Const HDI_FORMAT = &H4
Private Const HDI_TEXT = &H2
Private Const HDI_WIDTH = &H1
Private Const HDI_HEIGHT = HDI_WIDTH

Private Const HDF_LEFT = 0
Private Const HDF_RIGHT = 1
Private Const HDF_IMAGE = &H800
Private Const HDF_BITMAP_ON_RIGHT = &H1000
Private Const HDF_BITMAP = &H2000
Private Const HDF_STRING = &H4000

Private Const HDM_FIRST = &H1200
Private Const HDM_SETITEM = (HDM_FIRST + 4)

Private Const LVM_FIRST = &H1000
Private Const LVM_GETHEADER = (LVM_FIRST + 31)
Private Const LVM_GETBKCOLOR = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Private Const LVM_GETITEMA = (LVM_FIRST + 5)
Private Const LVM_GETITEM = LVM_GETITEMA
Private Const LVM_SETITEMA = (LVM_FIRST + 6)
Private Const LVM_SETITEM = LVM_SETITEMA
Private Const LVM_INSERTITEMA = (LVM_FIRST + 7)
Private Const LVM_INSERTITEM = LVM_INSERTITEMA
Private Const LVM_DELETEITEM = (LVM_FIRST + 8)
Private Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Private Const LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
Private Const LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Private Const LVM_FINDITEMA = (LVM_FIRST + 13)
Private Const LVM_FINDITEM = LVM_FINDITEMA
Private Const LVM_GETITEMRECT = (LVM_FIRST + 14)
Private Const LVM_SETITEMPOSITION = (LVM_FIRST + 15)
Private Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Private Const LVM_GETSTRINGWIDTHA = (LVM_FIRST + 17)
Private Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHA
Private Const LVM_HITTEST = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Private Const LVM_SCROLL = (LVM_FIRST + 20)
Private Const LVM_REDRAWITEMS = (LVM_FIRST + 21)
Private Const LVM_ARRANGE = (LVM_FIRST + 22)
Private Const LVM_EDITLABELA = (LVM_FIRST + 23)
Private Const LVM_EDITLABEL = LVM_EDITLABELA
Private Const LVM_GETEDITCONTROL = (LVM_FIRST + 24)
Private Const LVM_GETCOLUMNA = (LVM_FIRST + 25)
Private Const LVM_GETCOLUMN = LVM_GETCOLUMNA
Private Const LVM_SETCOLUMNA = (LVM_FIRST + 26)
Private Const LVM_SETCOLUMN = LVM_SETCOLUMNA
Private Const LVM_INSERTCOLUMNA = (LVM_FIRST + 27)
Private Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNA
Private Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
Private Const LVM_GETVIEWRECT = (LVM_FIRST + 34)
Private Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Private Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Private Const LVM_GETORIGIN = (LVM_FIRST + 41)
Private Const LVM_UPDATE = (LVM_FIRST + 42)
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXTA = (LVM_FIRST + 45)
Private Const LVM_GETITEMTEXT = LVM_GETITEMTEXTA
Private Const LVM_SETITEMTEXTA = (LVM_FIRST + 46)
Private Const LVM_SETITEMTEXT = LVM_SETITEMTEXTA
Private Const LVM_SETITEMCOUNT = (LVM_FIRST + 47)
Private Const LVM_SORTITEMS = (LVM_FIRST + 48)
Private Const LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Private Const LVM_GETITEMSPACING = (LVM_FIRST + 51)
Private Const LVM_GETISEARCHSTRINGA = (LVM_FIRST + 52)
Private Const LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGA
Private Const LVM_SETICONSPACING = (LVM_FIRST + 53)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Private Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
Private Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Private Const LVM_SETHOTITEM = (LVM_FIRST + 60)
Private Const LVM_GETHOTITEM = (LVM_FIRST + 61)
Private Const LVM_SETHOTCURSOR = (LVM_FIRST + 62)
Private Const LVM_GETHOTCURSOR = (LVM_FIRST + 63)
Private Const LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Type HD_ITEM
      mask        As Long
      cxy         As Long
      pszText     As String
      hbm         As Long
      cchTextMax  As Long
      fmt         As Long
      lParam      As Long
      iImage      As Long
      iOrder      As Long
End Type

Private Target As ListView

Private Const Format_Integer As String = "00000"
Private Const Format_Long As String = "0000000000"
Private Const Format_Decimal As String = "0000.0000"
Private Const Format_Date As String = "yyyy-mm-dd"
Private Const Format_Time As String = "hh-mm-ss"


Public Enum uSc_SortType
    [uSc_SortType_Normal] = 0
    [uSc_SortType_Title] = 1
    [uSc_SortType_Integer] = 2
    [uSc_SortType_Long] = 3
    [uSc_SortType_Decimal] = 4
    [uSc_SortType_Date] = 5
    [uSc_SortType_Time] = 6
End Enum

Public Enum uSc_SortOrder
    [uSc_SortOrder_Ascending] = 0
    [uSc_SortOrder_Descending] = 1
End Enum

Dim CurrentSortIndex As Integer
Dim CurrentSortType As uSc_SortType
Dim CurrentSortOrder As uSc_SortOrder




Public Sub Sort(ByVal Sorted As Boolean, Optional ByVal NewSortIndex As Integer = -1, Optional ByVal NewSortType As uSc_SortType)
Dim i As Long

If Not (Sorted) Then
    Target.Sorted = False
    Exit Sub
ElseIf NewSortIndex = CurrentSortIndex Then
    'Only change SortOrder
    Target.Sorted = False
    CurrentSortOrder = IIf(CurrentSortOrder = uSc_SortOrder_Ascending, uSc_SortOrder_Descending, uSc_SortOrder_Ascending)
    Target.SortOrder = CurrentSortOrder
    Target.Sorted = True
    SetHeaderIcons NewSortIndex
ElseIf NewSortIndex = -1 Then
    'Refresh
    NewSortIndex = CurrentSortIndex
    CurrentSortIndex = 0
    Sort True, NewSortIndex, CurrentSortType
Else
    CurrentSortIndex = NewSortIndex
    CurrentSortType = NewSortType
    CurrentSortOrder = uSc_SortOrder_Ascending
    Target.SortOrder = CurrentSortOrder
    
    Target.Sorted = False
    If CurrentSortType > 0 And CurrentSortType < 7 Then
        If CurrentSortIndex = 1 Then 'First col is not SubItem!
            For i = 1 To Target.ListItems.Count
                Target.ListItems(i).SubItems(Target.ColumnHeaders.Count - 1) = CFormat(Target.ListItems(i).Text, CurrentSortType)
            Next
        Else
            For i = 1 To Target.ListItems.Count
                Target.ListItems(i).SubItems(Target.ColumnHeaders.Count - 1) = CFormat(Target.ListItems(i).SubItems(CurrentSortIndex - 1), CurrentSortType)
            Next
        End If
        Target.SortKey = Target.ColumnHeaders.Count - 1
    Else
        Target.SortKey = CurrentSortIndex - 1
    End If
    Target.Sorted = True
    
SetHeaderIcons NewSortIndex
       
End If


End Sub


Private Function CFormat(Target As String, SortType As uSc_SortType) As String
Select Case SortType
'#### Title ######################################
    Case uSc_SortType_Title:
        CFormat = Trim(LCase(Target))
        If Left(CFormat, 4) = "the " Then CFormat = Right(CFormat, Len(CFormat) - 4)
        If Left(CFormat, 2) = "a " Then CFormat = Right(CFormat, Len(CFormat) - 2)
        'YOu can add prefixes like la, el, das, etc. here
        
'#### Integer ######################################
    Case uSc_SortType_Integer:
        If Val(Target) >= 0 Then
            CFormat = Format(Target, "+" & Format_Integer)
        Else
            CFormat = Format(Target, Format_Integer)
        End If
        
        'If you realy insist on sorting positive
        'and negative values you can use this instead:
        
         'CFormat = Format(Val(Target) + 32768, Format_Integer)
        
        '( + 2147483648 for Long Integer!)
        
'#### Long ######################################
    Case uSc_SortType_Long:
        If Val(Target) >= 0 Then
            CFormat = Format(Target, "+" & Format_Long)
        Else
            CFormat = Format(Target, Format_Integer)
        End If
        
'#### Decimal ######################################
    Case uSc_SortType_Decimal:
        If CSng(Target) >= 0 Then
            CFormat = Format(Target, "+" & Format_Decimal)
        Else
            CFormat = Format(Target, Format_Decimal)
        End If
        
'#### Date ######################################
    Case uSc_SortType_Date:
        CFormat = Format(Target, Format_Date)
'#### Time ######################################
    Case uSc_SortType_Time:
        CFormat = Format(Target, Format_Time)

'#### ... ######################################
    Case Else:
        CFormat = Target
    End Select
End Function


Public Sub SetUp(myListView As ListView)
    
    CurrentSortIndex = 0
    CurrentSortType = 0
    CurrentSortOrder = 0
    
    Set Target = myListView
    Sort True, 1, 1
    
End Sub




Public Sub SetHeaderIcons(SortBy As Integer)
Dim i As Integer
    For i = 1 To Target.ColumnHeaders.Count
     If i = SortBy Then
        Call ShowHeaderIcon(Target, i, HDF_IMAGE)
     Else
        Call ShowHeaderIcon(Target, i, 0)
     End If
   Next i
End Sub


Public Sub ShowHeaderIcon(ByVal myListView As ListView, ByVal colNumber As Long, ByVal showImage As Long)

Dim r As Long
Dim hHeader As Long
Dim HD As HD_ITEM

hHeader = SendMessageLong(myListView.hwnd, LVM_GETHEADER, 0, 0)

HD.mask = HDI_IMAGE Or HDI_FORMAT
HD.fmt = HDF_LEFT Or HDF_STRING Or HDF_BITMAP_ON_RIGHT Or showImage
HD.pszText = myListView.ColumnHeaders(myListView.SortKey + 1).Text

If showImage Then
    HD.iImage = myListView.SortOrder
End If

SendMessageAny hHeader, HDM_SETITEM, colNumber - 1, HD

End Sub
