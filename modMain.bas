Attribute VB_Name = "modMain"
Option Explicit
 
Public Declare Function DispCallFunc& Lib "oleaut32" (ByVal ppv&, ByVal oVft&, ByVal CC As Long, ByVal rtTYP%, ByVal paCount&, paTypes%, paValues&, fuReturn)
Public Declare Function OleCreatePropertyFrame& Lib "oleaut32" (ByVal hWndOwner&, ByVal X&, ByVal Y&, ByVal lpszCaption&, ByVal cObjects&, ByRef ppUnk&, ByVal cPages&, ByVal pPageClsID&, ByVal lcid&, ByVal dwReserved&, ByVal pvReserved&)

Public Const CLSID_ActiveMovieCategories = "{DA4E3DA0-D07D-11d0-BD50-00A0C911CE86}"
Public Const CLSID_VideoInputDeviceCategory = "{860BB310-5D01-11d0-BD3B-00A0C911CE86}"

'instantiates a new FilterGraphManager from a previously saved *.grf File (e.g. one, created with GraphEdit.exe)
Public Function CreateGraphFromFile(FileName As String) As FilgraphManager
Dim Stg As olelib.IStorage, Strm As olelib.IStream, PersStrm As olelib.IPersistStream
  If StgIsStorageFile(FileName) <> S_OK Then Err.Raise vbObjectError, , "Not a valid FilterGraph-File"

  Set Stg = StgOpenStorage(FileName, Stg, STGM_TRANSACTED Or STGM_READ Or STGM_SHARE_DENY_WRITE, vbNullString, 0)
   If Stg Is Nothing Then Err.Raise vbObjectError, , "Couldn't open Storage-File"

  Set Strm = Stg.OpenStream("ActiveMovieGraph", 0, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0)
   If Strm Is Nothing Then Err.Raise vbObjectError, , "Couldn't create ActiveMovieGraph-Stream"
 
  Set CreateGraphFromFile = New FilgraphManager
  Set PersStrm = CreateGraphFromFile
      PersStrm.Load Strm
End Function

'just the write-direction, to have a companion to the function above
Public Sub SaveGraphToFile(Graph As FilgraphManager, FileName As String)
Dim Stg As olelib.IStorage, Strm As olelib.IStream, PersStrm As olelib.IPersistStream
   Set Stg = StgCreateDocfile(FileName, STGM_CREATE Or STGM_TRANSACTED Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0)
    If Stg Is Nothing Then Err.Raise vbObjectError, , "Couldn't create new Storage-File"
 
   Set Strm = Stg.CreateStream("ActiveMovieGraph", STGM_CREATE Or STGM_WRITE Or STGM_SHARE_EXCLUSIVE, 0, 0)
    If Strm Is Nothing Then Err.Raise vbObjectError, , "Couldn't create new ActiveMovieGraph-Stream"
   
   Set PersStrm = Graph 'simple cast
       PersStrm.Save Strm, 1
   Stg.Commit STGC_DEFAULT
End Sub

'enumerates Group-Categories (a FriendlyNames-Array as the direct Function-Result, and a ClsID-Array with the accompanying ClsIDs of the Groups)
Public Function SubCategoriesOf(ByVal CategoryClsID As String, Optional SubCatClsIDs)
Const HKCR = &H80000000: Dim i, oReg, SubKeys, Result()
  
  SubCategoriesOf = Array() 'pre-initialize our return-value (so that a For Each becomes possible in either case)
  
  CategoryClsID = "CLSID\" & CategoryClsID & "\Instance"
  Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
      oReg.EnumKey HKCR, CategoryClsID, SubKeys
 
  If IsArray(SubKeys) Then
    ReDim Result(0 To UBound(SubKeys))
    If Not IsMissing(SubCatClsIDs) Then ReDim SubCatClsIDs(0 To UBound(SubKeys))
    
    For i = 0 To UBound(SubKeys)
      If Not IsMissing(SubCatClsIDs) Then oReg.GetStringValue HKCR, CategoryClsID & "\" & SubKeys(i), "CLSID", SubCatClsIDs(i)
      oReg.GetStringValue HKCR, CategoryClsID & "\" & SubKeys(i), "FriendlyName", Result(i)
    Next
    SubCategoriesOf = Result
  End If
End Function

'and here's the "go-one-level-deeper"-companion to the function above -> returns FriendlyNames again for the concrete Group -
'but the accompanying and related Container this time is not another Variant-Array, but a Collection with matching IBaseFilter-instances
Public Function GetDevicesInCategory(ByVal CatClsID As String, IBaseFiltersCol As Collection)
Const CLSID_SystemDeviceEnum = "{62BE5D10-60EB-11d0-BD3B-00A0C911CE86}"
Const IID_ICreateDevEnum = "{29840822-5b84-11d0-bd3b-00a0c911ce86}"
Const IID_IPropertyBag = "{55272A00-42CB-11CE-8135-00AA004BB851}"
Const IID_IBaseFilter = "{56A86895-0AD4-11CE-B03A-0020AF0BA770}"

Dim oUnk As stdole.IUnknown, Result()
Dim uuidCat As olelib.UUID, EnumMonikerCat As olelib.IEnumMoniker, Flags As Long
Dim CurMoniker As olelib.IMoniker, Fetched As Long
Dim uuidPBag As olelib.UUID, PropBag As olelib.IPropertyBag
Dim uuidBFlt As olelib.UUID, BaseFlt As IBaseFilter

  GetDevicesInCategory = Array() 'initialize an empty-array-result
  Set IBaseFiltersCol = New Collection
  
  Set oUnk = CreateInstanceUnk(CLSID_SystemDeviceEnum, IID_ICreateDevEnum)
  If oUnk Is Nothing Then Exit Function
  
  olelib.CLSIDFromString CatClsID, uuidCat
  vtblCall ObjPtr(oUnk), 3, VarPtr(uuidCat), VarPtr(EnumMonikerCat), Flags
  If EnumMonikerCat Is Nothing Then Exit Function

  olelib.CLSIDFromString IID_IPropertyBag, uuidPBag
  olelib.CLSIDFromString IID_IBaseFilter, uuidBFlt
  Result = Array()
  Do While EnumMonikerCat.Next(1, CurMoniker, Fetched) = S_OK
    On Error Resume Next
      CurMoniker.BindToObject Nothing, Nothing, uuidBFlt, BaseFlt
      CurMoniker.BindToStorage Nothing, Nothing, uuidPBag, PropBag
    On Error GoTo 0
    
    If Not BaseFlt Is Nothing And Not PropBag Is Nothing Then
       ReDim Preserve Result(0 To UBound(Result) + 1)
       PropBag.Read StrPtr("FriendlyName"), Result(UBound(Result)), Nothing
       IBaseFiltersCol.Add BaseFlt
    End If
    
    Set BaseFlt = Nothing
    Set PropBag = Nothing
    Set CurMoniker = Nothing
  Loop
  GetDevicesInCategory = Result
End Function

'a Helper-Function, to instantiate comfortably per "String-based-ClsID/IID"
Public Function CreateInstanceUnk(ClsID As String, IID As String) As stdole.IUnknown
Static uuidCls As olelib.UUID, uuidIID As olelib.UUID, HResult As Long
  If Len(ClsID) <> 38 Or Len(IID) <> 38 Then Exit Function
  
  olelib.CLSIDFromString ClsID, uuidCls
  olelib.CLSIDFromString IID, uuidIID
 
  HResult = olelib.CoCreateInstance(uuidCls, Nothing, CLSCTX_INPROC_SERVER, uuidIID, CreateInstanceUnk)
  If HResult <> S_OK Then Err.Raise HResult
End Function
'another Helper, somewhat similar to the above, but for casts by IIDString from existing instances
Public Function CastToUnkByIID(ByVal ObjToCastFrom As olelib.IUnknown, IID As String) As stdole.IUnknown
Dim UUID As olelib.UUID
  olelib.CLSIDFromString IID, UUID
  ObjToCastFrom.QueryInterface UUID, CastToUnkByIID
End Function

'and that's a helper in case one wants to be lazy and not write a TypeLib explicitely ... useful,
'when the interfaces in question has only a few members - or is in other ways difficult with regards
'to the passed parameters... this routine allows a few more degrees of "creativity" with the arguments.
Public Function vtblCall(ByVal pUnk As Long, ByVal vtblIdx As Long, ParamArray P() As Variant)
Static VType(0 To 31) As Integer, VPtr(0 To 31) As Long
Dim i As Long, V(), HResDisp As Long
  If pUnk = 0 Then vtblCall = 5: Exit Function

  V = P 'make a copy of the params, to prevent problems with VT_ByRef-Members in the ParamArray
  For i = 0 To UBound(V)
    VType(i) = VarType(V(i))
    VPtr(i) = VarPtr(V(i))
  Next i
  
  HResDisp = DispCallFunc(pUnk, vtblIdx * 4, 4, vbLong, i, VType(0), VPtr(0), vtblCall)
  If HResDisp <> S_OK Then Err.Raise HResDisp, , "Error in DispCallFunc"
End Function
  
Public Function BStrFromLPWStr(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
  SysReAllocString BStrFromLPWStr, lpWStr
  If CleanupLPWStr Then CoTaskMemFree lpWStr
End Function

Public Function CheckForFileSinkAndSetFileName(ByVal Flt As olelib.IUnknown, FileName As String) As Boolean
Const IID_IFileSinkFilter = "{A2104830-7C70-11CF-8BCE-00AA00A3F1A6}", VTbl_SetFileName = 3
Dim oUnkFSink As stdole.IUnknown

  Set oUnkFSink = CastToUnkByIID(Flt, IID_IFileSinkFilter)
  CheckForFileSinkAndSetFileName = vtblCall(ObjPtr(oUnkFSink), VTbl_SetFileName, StrPtr(FileName), 0&) = S_OK
End Function
Public Function CheckForFileSinkAndGetFileName(ByVal Flt As olelib.IUnknown, FileName As String) As Boolean
Const IID_IFileSinkFilter = "{A2104830-7C70-11CF-8BCE-00AA00A3F1A6}", VTbl_GetFileName = 4
Dim oUnkFSink As stdole.IUnknown, lpWFileName As Long

  Set oUnkFSink = CastToUnkByIID(Flt, IID_IFileSinkFilter)
  If vtblCall(ObjPtr(oUnkFSink), VTbl_GetFileName, VarPtr(lpWFileName), 0&) Then Exit Function
  
  FileName = BStrFromLPWStr(lpWFileName) 'hand out the ByRef-Argument (and cleanup/free lpWFileName)
  CheckForFileSinkAndGetFileName = True
End Function

Public Function ShowPropertyPage(ByVal FilterOrPin As olelib.IUnknown, Optional Caption As String, Optional ByVal hWndOwner As Long) As Boolean
Const IID_ISpecifyPropertyPages = "{B196B28B-BAB4-101A-B69C-00AA00341D07}", VTbl_GetPages = 3
Dim oUnkSpPP As stdole.IUnknown, CAUUID(0 To 1) As Long
  Set oUnkSpPP = CastToUnkByIID(FilterOrPin, IID_ISpecifyPropertyPages)
 
  If vtblCall(ObjPtr(oUnkSpPP), VTbl_GetPages, VarPtr(CAUUID(0))) Then Exit Function
  If CAUUID(0) = 0 Then Exit Function 'no PropPageCount was returned
  
  OleCreatePropertyFrame hWndOwner, 0, 0, StrPtr(Caption), 1, ObjPtr(FilterOrPin), CAUUID(0), CAUUID(1), 0, 0, 0

  CoTaskMemFree CAUUID(1)
  ShowPropertyPage = True
End Function
