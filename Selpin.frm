VERSION 5.00
Begin VB.Form frmSelectPin 
   Caption         =   "Connect to Pin"
   ClientHeight    =   3210
   ClientLeft      =   4890
   ClientTop       =   4920
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3210
   ScaleWidth      =   6270
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox listPins 
      Height          =   1425
      ItemData        =   "Selpin.frx":0000
      Left            =   3360
      List            =   "Selpin.frx":0002
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.ListBox listFilters 
      Height          =   1425
      ItemData        =   "Selpin.frx":0004
      Left            =   240
      List            =   "Selpin.frx":0006
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Pins"
      Height          =   252
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   492
   End
   Begin VB.Label VendorInfoLabel 
      Caption         =   "Vendor Info:"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label VendorInfo 
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Filters"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmSelectPin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_objFI As IFilterInfo, g_objPI As IPinInfo, g_objMC As IMediaControl

Public OtherDir As Long, bOK As Boolean

Private Sub Form_Load()
Dim Pin As IPinInfo, Filter As IFilterInfo, pinOther As IPinInfo
    On Error Resume Next
     
    listFilters.Clear
    For Each Filter In g_objMC.FilterCollection
        For Each Pin In Filter.Pins
            Set pinOther = Pin.ConnectedTo
            If Err Then
               Err.Clear
               If Pin.Direction <> OtherDir And g_objFI.Name <> Filter.Name Then
                  listFilters.AddItem Filter.Name
                  Exit For
               End If
            End If
        Next
    Next

    If listFilters.ListCount > 0 Then listFilters.ListIndex = 0 'reset the list index
End Sub
 
Private Sub OK_Click()
    bOK = Not g_objPI Is Nothing
    Unload Me
End Sub
 
Private Sub Cancel_Click()
    bOK = False
    Unload Me
End Sub

Private Sub listFilters_Click()
Dim Pin As IPinInfo, pfilter As IFilterInfo, pinOther As IPinInfo
    On Error Resume Next
    
    'enumerate through each filter in the filter collection
    For Each pfilter In g_objMC.FilterCollection
        If LCase(pfilter.Name) = LCase(listFilters.Text) Then
            Set g_objFI = pfilter ' global FilterInfo object
            VendorInfo.Caption = pfilter.VendorInfo
            
            listPins.Clear
            Set g_objPI = Nothing
            For Each Pin In pfilter.Pins 'enumerate through each pin in the filter
                Set pinOther = Pin.ConnectedTo
                If Err Then
                  Err.Clear
                  If Pin.Direction <> OtherDir Then listPins.AddItem Pin.Name
                End If
            Next
        End If
    Next

    If listPins.ListCount > 0 Then listPins.ListIndex = 0 'reset the selected index
End Sub

Private Sub listPins_Click() 'a new pin is selected, store it in the module-level pin object
Dim objPinInfo As IPinInfo
  Set g_objPI = Nothing
  If listPins.ListIndex < 0 Then Exit Sub
  
  On Error Resume Next
  For Each objPinInfo In g_objFI.Pins 'enumerate the pins
    If objPinInfo.Name = listPins.Text Then Set g_objPI = objPinInfo
  Next
End Sub
