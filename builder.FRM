VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Filter and Pin Viewer"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   16695
   Icon            =   "builder.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1113
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picPrev 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   5835
      Left            =   8790
      ScaleHeight     =   389
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   30
      Top             =   150
      Width           =   7695
   End
   Begin VB.CommandButton btnPinProps 
      Caption         =   "Pin Properties"
      Height          =   375
      Left            =   2850
      TabIndex        =   28
      Top             =   6120
      Width           =   1395
   End
   Begin MSComDlg.CommonDialog ctrlCommonDlg 
      Left            =   3750
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   4096
   End
   Begin VB.CommandButton btnVideoControl 
      Caption         =   "IAMVideoControl"
      Height          =   375
      Left            =   6330
      TabIndex        =   27
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton btnTestCam 
      Caption         =   "Test Camera"
      Height          =   375
      Left            =   4980
      TabIndex        =   26
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save *.grf"
      Height          =   255
      Left            =   7500
      TabIndex        =   25
      Top             =   300
      Width           =   885
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load *.grf"
      Height          =   255
      Left            =   6510
      TabIndex        =   24
      Top             =   300
      Width           =   915
   End
   Begin VB.CommandButton x2 
      Caption         =   "2x Speed"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton btnFilterProps 
      Caption         =   "Filter Properties"
      Height          =   375
      Left            =   1320
      TabIndex        =   21
      Top             =   6120
      Width           =   1395
   End
   Begin VB.Frame frameFilter 
      Caption         =   "Filter Graph"
      Height          =   2652
      Left            =   75
      TabIndex        =   18
      Top             =   75
      Width           =   8565
      Begin VB.CommandButton Remove 
         Caption         =   "Remove"
         Height          =   315
         Left            =   4320
         TabIndex        =   29
         Top             =   2190
         Width           =   915
      End
      Begin VB.CommandButton cmdCategory 
         Caption         =   "Add by Categ..."
         Height          =   315
         Left            =   6960
         TabIndex        =   22
         Top             =   2190
         Width           =   1335
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   2175
         TabIndex        =   3
         Top             =   2175
         Width           =   840
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   2175
         Width           =   840
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   315
         Left            =   225
         TabIndex        =   1
         Top             =   2175
         Width           =   840
      End
      Begin VB.CommandButton cmdSource 
         Caption         =   "Add &Source Filter..."
         Height          =   315
         Left            =   5310
         TabIndex        =   6
         Top             =   2190
         Width           =   1572
      End
      Begin VB.ListBox lstFilters 
         Height          =   1620
         ItemData        =   "builder.frx":0442
         Left            =   4320
         List            =   "builder.frx":0449
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   510
         Width           =   4005
      End
      Begin VB.CommandButton cmdAddRegFilter 
         Caption         =   "&Add ->"
         Height          =   315
         Left            =   3660
         TabIndex        =   4
         Top             =   480
         Width           =   645
      End
      Begin VB.ListBox lstRegFilters 
         Height          =   1620
         ItemData        =   "builder.frx":0459
         Left            =   210
         List            =   "builder.frx":0460
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   3405
      End
      Begin VB.Label lblFilters 
         Caption         =   "Filters in current filter &graph"
         Height          =   255
         Left            =   4290
         TabIndex        =   13
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label lblRegFilters 
         Caption         =   "&Registered filters"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame framePinInfo 
      Caption         =   "Filter"
      Height          =   3132
      Left            =   75
      TabIndex        =   11
      Top             =   2850
      Width           =   8565
      Begin VB.TextBox txtPinInfo 
         Height          =   1620
         Left            =   4320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   10
         Top             =   1200
         Width           =   4005
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Co&nnect One Pin..."
         Height          =   315
         Left            =   2025
         TabIndex        =   9
         Top             =   2500
         Width           =   1575
      End
      Begin VB.CommandButton cmdRender 
         Caption         =   "&Connect Downstream"
         Height          =   315
         Left            =   225
         TabIndex        =   8
         Top             =   2500
         Width           =   1692
      End
      Begin VB.ListBox lstPins 
         Height          =   1230
         ItemData        =   "builder.frx":0473
         Left            =   210
         List            =   "builder.frx":047A
         TabIndex        =   7
         Top             =   1200
         Width           =   3405
      End
      Begin VB.Label lblFilterName 
         Height          =   255
         Left            =   1245
         TabIndex        =   20
         Top             =   225
         Width           =   4995
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter name:"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblVendor 
         Caption         =   "Vendor: "
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblVendorInfo 
         Height          =   255
         Left            =   1245
         TabIndex        =   16
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label lblPinListbox 
         Caption         =   "&Pins in selected filter"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   940
         Width           =   2055
      End
      Begin VB.Label lblPinInfo 
         Caption         =   "Information for selected pin"
         Height          =   285
         Left            =   4380
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Menu mnuFilterGraph 
      Caption         =   "&FilterGraph"
      Begin VB.Menu mnu_FilterGraphNew 
         Caption         =   "&New (empty)"
      End
      Begin VB.Menu mnu_FilterGraphGenerate 
         Caption         =   "&Generate from input file..."
      End
      Begin VB.Menu mnu_Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_FilterGraphRun 
         Caption         =   "&Run"
      End
      Begin VB.Menu mnu_FilterGraphPause 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnu_FilterGraphStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnu_Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_FilterGraphExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_BuildCustomGraph 
         Caption         =   "&Build custom graph"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private m_GraphManager As IMediaControl   'IMediaControl interface provided by IFilgraphManager
Private m_CurSelFilter As IFilterInfo, m_CurSelPin As IPinInfo  'set in the appropr. ListBox-Events

Private Sub Form_Load()
  InitializeNew
End Sub

Private Sub InitializeNew()
  Set m_GraphManager = New FilgraphManager
  Set m_CurSelFilter = Nothing
  Set m_CurSelPin = Nothing
  
  lstFilters.Clear: lstPins.Clear: txtPinInfo.Text = "" 'reset GUI-Controls
  lblFilterName.Caption = "": lblVendorInfo.Caption = ""
  
  RefreshRegFilters ' update the display for registered filters
End Sub

Private Property Get GraphState(Optional msTimeOut As Long) As [_FilterState] '<- this type is provided by msvidctl.dll
  If Not m_GraphManager Is Nothing Then m_GraphManager.GetState msTimeOut, GraphState
End Property
 
' appends a filter to the current working graph
Private Sub AppendFilter(bstrFilterName As String, objFilter As IFilterInfo)
Dim objRegFilter As IRegFilterInfo
On Error GoTo ErrLine
 
    'step through each registered filter for verification and proceed to append the filter if the filter is in fact valid..
    For Each objRegFilter In m_GraphManager.RegFilterCollection
        If StrComp(objRegFilter.Name, bstrFilterName, vbTextCompare) = 0 Then
            objRegFilter.Filter objFilter
            Exit For
        End If
    Next

    RefreshFilters
ErrLine:
End Sub
            
'update the listbox of registered filters (listbox lstRegFilters)
Private Sub RefreshRegFilters()
Dim objRegFilter As IRegFilterInfo
On Error GoTo ErrLine
 
    lstRegFilters.Clear
    
    For Each objRegFilter In m_GraphManager.RegFilterCollection
        lstRegFilters.AddItem objRegFilter.Name
        lstRegFilters.ItemData(lstRegFilters.NewIndex) = lstRegFilters.ListCount - 1
    Next
 
    If lstRegFilters.ListCount Then lstRegFilters.ListIndex = 0    ' select first in list
ErrLine:
End Sub

Private Sub btnLoad_Click()
Dim NewGraph As FilgraphManager
    If Not ShowFileDialog("Graph files (*.grf)|*.grf|") Then Exit Sub
    
    mnu_FilterGraphNew_Click
 
    On Error Resume Next
      Set NewGraph = CreateGraphFromFile(ctrlCommonDlg.FileName)
    On Error GoTo 0
    If Not NewGraph Is Nothing Then Set m_GraphManager = NewGraph
    RefreshFilters
End Sub

Private Sub btnSave_Click()
    If ShowFileDialog("Graph files (*.grf)|*.grf|", True) Then SaveGraphToFile m_GraphManager, ctrlCommonDlg.FileName
End Sub
 
Private Sub btnFilterProps_Click()
  If Not m_CurSelFilter Is Nothing Then ShowPropertyPage m_CurSelFilter.Filter, m_CurSelFilter.Name, Me.hWnd
End Sub
Private Sub btnPinProps_Click()
  If Not m_CurSelPin Is Nothing Then ShowPropertyPage m_CurSelPin.Pin, m_CurSelPin.Name, Me.hWnd
End Sub

Private Sub Remove_Click()
Dim GraphBuilder As IGraphBuilder: Set GraphBuilder = m_GraphManager
    On Error Resume Next
        GraphBuilder.RemoveFilter m_CurSelFilter.Filter
    On Error GoTo 0
    lstPins.Clear: txtPinInfo.Text = "" 'reset GUI-Controls
    lblFilterName.Caption = "": lblVendorInfo.Caption = ""
    RefreshFilters
End Sub


Private Sub x2_Click()
Dim objSeek As IMediaPosition
    Set objSeek = m_GraphManager
    On Error Resume Next
      objSeek.Rate = 2
    If Err Then Err.Clear
    objSeek.CurrentPosition = 0
End Sub

' when the user manually appends a filter to the graph from a category-filtered list
Private Sub cmdCategory_Click()
    If m_GraphManager Is Nothing Then mnu_FilterGraphNew_Click

    frmCategory.Show 1
    If frmCategory.CurSelFilter Is Nothing Then Exit Sub
    
    Dim GraphBuilder As IGraphBuilder: Set GraphBuilder = m_GraphManager 'just a cast to a Type-Interface, defined in msvidctl.dll
    GraphBuilder.AddFilter frmCategory.CurSelFilter, frmCategory.CurSelFriendlyName
    RefreshFilters
End Sub

'just a delegation to the handler-function below
Private Sub lstRegFilters_DblClick()
    cmdAddRegFilter_Click
End Sub
' when the user manually appends a filter to the graph
Private Sub cmdAddRegFilter_Click()
Dim objFilter As IFilterInfo, objRegFilter As IRegFilterInfo, FileName As String
    On Error GoTo ErrLine
            
    If m_GraphManager Is Nothing Or lstRegFilters.ListIndex < 0 Then Exit Sub
 
    m_GraphManager.RegFilterCollection.Item lstRegFilters.ItemData(lstRegFilters.ListIndex), objRegFilter
    objRegFilter.Filter objFilter
     
    If objFilter.IsFileSource Then
       If ShowFileDialog() Then objFilter.FileName = ctrlCommonDlg.FileName
    End If

    If CheckForFileSinkAndGetFileName(objFilter.Filter, FileName) Then 'we have a filesink-Filter
      If Len(FileName) = 0 Then 'but its filename-property was not yet set
        If ShowFileDialog(, True) Then CheckForFileSinkAndSetFileName objFilter.Filter, ctrlCommonDlg.FileName
      End If
    End If
 
    RefreshFilters
ErrLine:
End Sub

'connects the current selected pin, indicated by m_CurSelPin , m_GraphManager, and m_CurSelFilter
Private Sub cmdConnect_Click()
On Error GoTo ErrLine ' if already connected, connect fails

    'set direction
    frmSelectPin.OtherDir = m_CurSelPin.Direction
    Set frmSelectPin.g_objMC = m_GraphManager
    Set frmSelectPin.g_objFI = m_CurSelFilter
    
    frmSelectPin.Show 1
    If Not frmSelectPin.bOK Then Exit Sub
    
    If Not frmSelectPin.g_objPI Is Nothing Then m_CurSelPin.Connect frmSelectPin.g_objPI
        
    RefreshFilters  ' refresh the filters in the ui
ErrLine:
End Sub
       
'when the user opts to connect the current graph, causing downstream pins to render
Private Sub cmdRender_Click()
On Error GoTo ErrLine
    m_CurSelPin.Render
    RefreshFilters
ErrLine:
End Sub

'Appends a source filter to the graph that can read the given filename.
Private Sub cmdSource_Click()
Dim i As Long, objFilter As Object
On Error GoTo ErrLine
 
    If Not ShowFileDialog("Media files (*.mpg;*.avi;*.mov)|*.mpg;*.avi;*.mov|") Then Exit Sub
    
    'verify that the filter has not already been appended to the list
    For i = 0 To Me.lstFilters.ListCount
      If StrComp(lstFilters.List(i), ctrlCommonDlg.FileName, vbTextCompare) = 0 Then Exit Sub
    Next
    Call m_GraphManager.AddSourceFilter(ctrlCommonDlg.FileName, objFilter)
    RefreshFilters  ' update all info displayed by this VB app
ErrLine:
End Sub
            
'Refresh the contents of the "Filters" combo box using the current IMediaControl interface
Public Sub RefreshFilters()
Dim objFilter As IFilterInfo
On Error GoTo ErrLine
 
    lstFilters.Clear
    For Each objFilter In m_GraphManager.FilterCollection
        lstFilters.AddItem objFilter.Name  'append the filter's name
        lstFilters.ItemData(lstFilters.NewIndex) = lstFilters.ListCount - 1
    Next
    If lstFilters.ListCount Then lstFilters.ListIndex = 0  'select first in list
ErrLine:
End Sub
                     
Private Sub lstFilters_Click()
Dim objPinInfo As IPinInfo
On Error GoTo ErrLine
    Set m_CurSelFilter = Nothing
    lstPins.Clear
    If lstFilters.ListIndex < 0 Then Exit Sub

    m_GraphManager.FilterCollection.Item lstFilters.ItemData(lstFilters.ListIndex), m_CurSelFilter
    lblFilterName.Caption = m_CurSelFilter.Name
    lblVendorInfo.Caption = m_CurSelFilter.VendorInfo

    For Each objPinInfo In m_CurSelFilter.Pins
      lstPins.AddItem objPinInfo.Name
    Next

    If lstPins.ListCount Then lstPins.ListIndex = 0 ' select the first in the list
ErrLine:
End Sub
            
'Proceed to Append detailed information on the selected pin to the textbox
'This procedure is designed whereas it will encapsulate the error if the pin
'cannot be connectedto, and still set the correct text.  This is by design.
Private Sub lstPins_Click()
Dim strSnippet As String, objPinInfo As IPinInfo, objPeerFilter As IFilterInfo
Dim objConnPinInfo As IPinInfo, objMediaTypeInfo As IMediaTypeInfo
On Error Resume Next
 
    For Each objPinInfo In m_CurSelFilter.Pins 'enumerate throuch each pin
        If StrComp(objPinInfo.Name, lstPins.Text, vbTextCompare) = 0 Then
               
               Set m_CurSelPin = objPinInfo 'reset the selected pin
               
               'attempt to obtain a reference to the connected pin; always anticipating that an error could occure..
               Set objConnPinInfo = objPinInfo.ConnectedTo
               
               If Not objConnPinInfo Is Nothing Then 'the above method completed without error, list the pin info..
                 strSnippet = "Connected to pin: " & objConnPinInfo.Name & " "
                 If Not objConnPinInfo Is Nothing Then
                    If Not objConnPinInfo.FilterInfo Is Nothing Then
                       Set objPeerFilter = objConnPinInfo.FilterInfo
                       strSnippet = strSnippet & " on filter: " & objPeerFilter.Name & " "
                       Set objMediaTypeInfo = objPinInfo.ConnectionMediaType
                       strSnippet = strSnippet & vbCrLf & "Media Type: " & objMediaTypeInfo.Type
                    End If
                 End If
               End If
               
             'append info based upon pin direction..
             strSnippet = strSnippet & " " & vbCrLf & IIf(objPinInfo.Direction = 0, "Direction: Input", "Direction: Output")
 
             txtPinInfo.Text = strSnippet 'set the information to the textbox..
        End If
    Next
End Sub
        
Private Sub cmdPause_Click()
    mnu_FilterGraphPause_Click
End Sub
Private Sub cmdRun_Click()
    mnu_FilterGraphRun_Click
End Sub
Private Sub cmdStop_Click()
    mnu_FilterGraphStop_Click
End Sub

Private Sub mnu_FilterGraphExit_Click()
Dim frm As Form
On Error GoTo ErrLine
    For Each frm In Forms 'exit application
       frm.Visible = False: Unload frm
    Next
ErrLine:
End Sub
            
'User is initializing the filter graph based on a source file
Private Sub mnu_FilterGraphGenerate_Click()
On Error GoTo ErrLine
    InitializeNew
    
    If Not ShowFileDialog("Media files (*.mpg;*.avi;*.mov;*.wav)|*.mpg;*.avi;*.mov;*.wav|") Then Exit Sub
 
    'call IMediaControl::RenderFile to add all filters and connect all pins
    m_GraphManager.RenderFile ctrlCommonDlg.FileName
    RefreshFilters  'update the ui
ErrLine:
End Sub
            
Private Sub mnu_FilterGraphNew_Click()
    InitializeNew
End Sub
 
Private Sub mnu_FilterGraphRun_Click()
Dim objPosition As IMediaPosition
    If m_GraphManager Is Nothing Then Exit Sub

    On Error Resume Next
      Set objPosition = m_GraphManager
      objPosition.CurrentPosition = 0
      m_GraphManager.Run
      
      SetLocalPreviewBoxTo picPrev, m_GraphManager
    If Err Then Err.Clear
End Sub
Private Sub mnu_FilterGraphPause_Click()
On Error GoTo ErrLine
    If Not GraphState = State_Stopped Then m_GraphManager.Pause
ErrLine:
End Sub
Private Sub mnu_FilterGraphStop_Click()
On Error GoTo ErrLine
    If Not m_GraphManager Is Nothing Then m_GraphManager.Stop
ErrLine:
End Sub

Private Sub mnu_BuildCustomGraph_Click()
  MsgBox "Todo: Come up with a nice (and working) example which connects a few filters 'by-hand'"
End Sub

Private Function ShowFileDialog(Optional Filter As String, Optional ByVal ShowSave As Boolean) As Boolean
    ctrlCommonDlg.CancelError = False
    ctrlCommonDlg.Filter = Filter
    If ShowSave Then ctrlCommonDlg.ShowSave Else ctrlCommonDlg.ShowOpen
    ShowFileDialog = Len(ctrlCommonDlg.FileName)
End Function


Private Sub btnTestCam_Click()
MsgBox "Todo: VideoProcAmpInfo and CameraControlInfo"
'On Local Error GoTo ErrLine
'    Dim vpa As VideoProcAmpInfo
'    Set vpa = New VideoProcAmpInfo
'    vpa.SetFilter m_CurSelFilter
'    vpa.Set 0, 20, 0    ' brightness = 20
'    vpa.Set 1, 20, 0    ' contrast = 20
'
'    Dim cam As CameraControlInfo
'    Set cam = New CameraControlInfo
'    cam.SetFilter m_CurSelFilter
'    cam.Set 4, 40, 0    ' zoom = 40
'ErrLine:
End Sub

Private Sub btnVideoControl_Click()
MsgBox "Todo: VBVideoControl"
'On Local Error GoTo ErrLine
'    Dim vc As VBVideoControl
'    Set vc = New VBVideoControl
'    vc.SetPin m_CurSelPin
'    Dim res As Long
'    res = vc.Caps
'    res = vc.Mode
'ErrLine:
End Sub

Private Sub SetLocalPreviewBoxTo(PB As PictureBox, ByVal GraphVW As IVideoWindow)
  PB.ScaleMode = vbPixels
  On Error Resume Next
    GraphVW.Visible = False
    GraphVW.WindowStyle = &H44000000 '<- WS_CHILD Or WS_CLIPSIBLINGS)
    GraphVW.Owner = PB.hWnd
    GraphVW.FullScreenMode = False
    GraphVW.SetWindowPosition 0, 0, PB.ScaleWidth, PB.ScaleHeight
    GraphVW.Visible = True
  If Err Then Err.Clear
End Sub

