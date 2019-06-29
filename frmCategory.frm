VERSION 5.00
Begin VB.Form frmCategory 
   Caption         =   "Filters By Category"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   4500
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.ListBox listFilters 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox comboCategory 
      Height          =   315
      ItemData        =   "frmCategory.frx":0000
      Left            =   120
      List            =   "frmCategory.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurSelFriendlyName As String, CurSelFilter As IBaseFilter

Private CatGrpsFriendlyNames, CatGrpsClsIDs
Private FiltersFriendlyNames, BaseFiltersCol As Collection

Private Sub Form_Load()
Dim i As Long, IdxVidCaptureGroup As Long
   CatGrpsFriendlyNames = SubCategoriesOf(CLSID_ActiveMovieCategories, CatGrpsClsIDs)
   For i = 0 To UBound(CatGrpsFriendlyNames)
      If CatGrpsClsIDs(i) = CLSID_VideoInputDeviceCategory Then
         IdxVidCaptureGroup = i
      Else 'add into the sorting combo
        comboCategory.AddItem CatGrpsFriendlyNames(i)
        comboCategory.ItemData(comboCategory.NewIndex) = i
      End If
   Next i
   'just ensure, that we add the CLSID_VideoInputDeviceCategory last (using our stored IdxVidCaptureGroup-Index)
   comboCategory.AddItem CatGrpsFriendlyNames(IdxVidCaptureGroup)
   comboCategory.ItemData(comboCategory.NewIndex) = IdxVidCaptureGroup
   comboCategory.ListIndex = comboCategory.NewIndex
End Sub
 
Private Sub cmdCancel_Click()
    Set CurSelFilter = Nothing
    Unload Me
End Sub

Private Sub comboCategory_Click()
Dim FN, GrpClsID As String
    Set CurSelFilter = Nothing
    GrpClsID = CatGrpsClsIDs(comboCategory.ItemData(comboCategory.ListIndex))
    FiltersFriendlyNames = GetDevicesInCategory(GrpClsID, BaseFiltersCol)
        
    listFilters.Clear
    For Each FN In FiltersFriendlyNames
      listFilters.AddItem FN
      listFilters.ItemData(listFilters.NewIndex) = listFilters.ListCount - 1
    Next
    If listFilters.ListCount Then listFilters.ListIndex = 0
End Sub

Private Sub listFilters_Click()
  If listFilters.ListIndex < 0 Then Exit Sub
  CurSelFriendlyName = FiltersFriendlyNames(listFilters.ItemData(listFilters.ListIndex))
  Set CurSelFilter = BaseFiltersCol(listFilters.ItemData(listFilters.ListIndex) + 1)
End Sub

Private Sub listFilters_DblClick()
    cmdInsert_Click
End Sub
Private Sub cmdInsert_Click()
    Unload Me
End Sub
 
