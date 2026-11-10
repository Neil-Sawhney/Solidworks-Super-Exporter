'===============================================================================
' SWExporter - standard module.  Paste into a new module (Insert > Module).
' Successor to Lenny Kikstra's DocExport (v1.30).  Entry point: main
' Exports bodies, configurations, configs x bodies, assembly components,
' visible bodies (part or assembly), drawing sheets, or a batch of files to
' STL/STEP/IGES/X_T/SAT/3MF/GLB/OBJ/PDF/eDrawings/DWG/DXF/JPG/PNG.
' References (Tools > References):
'   SldWorks <year> Type Library
'   SOLIDWORKS <year> Constant type library
'   Microsoft Forms 2.0 Object Library (added automatically with a UserForm)
' Needs a UserForm named frmSWExporter (see the form code). SOLIDWORKS 2022+ (VBA7, 64-bit).
'===============================================================================
Option Explicit

' ---- Windows file dialog (multi-select) ----
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type
Private Declare PtrSafe Function GetOpenFileNameA Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

Public swApp As SldWorks.SldWorks

' ---- mode ids ----
Public Const MODE_BODIES As Long = 0
Public Const MODE_CONFIGS As Long = 1
Public Const MODE_CFGBODIES As Long = 2
Public Const MODE_SHEETS As Long = 3
Public Const MODE_BATCH As Long = 4
Public Const MODE_COMPONENTS As Long = 5
Public Const MODE_COMPBODIES As Long = 6
Public Const MODE_VISBODIES As Long = 7     ' visible solid bodies, one file each (part or assembly)

' ---- format table (filled from the ini or the built-in defaults) ----
Public Type ExportFormat
    Label As String          ' shown in the form
    Ext As String            ' ".stl", ".e*" ...
    Kind As String           ' "MODEL" or "DRAWING"
    Flag As String           ' MESH / ISOLATE / WHOLE / SINGLE / MULTI
    DefaultOn As Boolean
    Selected As Boolean
End Type
Public gFormats() As ExportFormat
Public gFormatCount As Long

' ---- options set by the form ----
Public gMode As Long
Public gSelItems As Collection       ' ticked configs / bodies / sheets / part keys / file paths
Public gIncludeSurfaces As Boolean
Public gOutMode As Long              ' 0 = next to file, 1 = subfolder, 2 = custom
Public gOutText As String            ' subfolder template or custom path
Public gTemplate As String           ' "" = automatic name per mode
Public gCollision As Long            ' 0 = auto-number, 1 = skip, 2 = overwrite
Public gStlUnits As Long             ' swLengthUnit_e: 0 mm, 1 cm, 2 m, 3 in, 4 ft (also used for OBJ)
Public gStlQuality As Long           ' swSTLQuality_e
Public gStlBinary As Boolean
Public gPdfCombined As Boolean
Public gDryRun As Boolean
Public gWriteLog As Boolean
Public gBatchFolder As String
Public gBatchModelMode As Long       ' MODE_CONFIGS / MODE_BODIES / MODE_CFGBODIES
Public gCfgFilter As String          ' batch: only configs containing this ("" = all)
Public gCfgFilterCase As Boolean
Public gCancel As Boolean
Public gFormLoaded As Boolean

' ---- run state ----
Private mLog As String
Private mFirstOutDir As String
Private mFirstSrcDir As String
Private mPlanned As Collection
Private mOk As Long, mFail As Long, mSkip As Long, mStep As Long
Private mRunStamp As String
Private mModelMode As Long           ' mode of the document being exported right now
Private mAnchorDoc As SldWorks.ModelDoc2   ' assembly whose folder/name is used while exporting its parts
Private mCompName As String          ' value of {component}
Private mDone As Object              ' parts already exported in this run (batch)
Private mFso As Object
Private mBodyFilter As Collection    ' assembly export: body names shown in the assembly (Nothing = no filter)
Private mFilterStrict As Boolean     ' True = never fall back to "all bodies" when no name matches

' ---- mesh buffer for the GLB/OBJ writer (meters, 9 floats per triangle) ----
Private mPos() As Single
Private mNrm() As Single
Private mCnt As Long

'===============================================================================
Sub main()
    Set swApp = Application.SldWorks
    LoadFormats
    frmSWExporter.Show
End Sub

Private Function FS() As Object
    If mFso Is Nothing Then Set mFso = CreateObject("Scripting.FileSystemObject")
    Set FS = mFso
End Function

'================================ INI ==========================================
Public Function MacroFolder() As String
    Dim p As String
    p = swApp.GetCurrentMacroPathName
    If InStrRev(p, "\") > 0 Then MacroFolder = Left$(p, InStrRev(p, "\") - 1)
End Function

Public Sub LoadFormats()
    Dim iniPath As String, ts As Object, ln As String, section As String
    gFormatCount = 0
    ReDim gFormats(0 To 31)
    On Error GoTo Fallback
    iniPath = MacroFolder() & "\SWExporter.ini"
    If FS().FileExists(iniPath) Then
        Set ts = FS().OpenTextFile(iniPath, 1)
        Do While Not ts.AtEndOfStream
            ln = Trim$(Replace(ts.ReadLine, vbTab, " "))
            If ln = "" Or Left$(ln, 1) = ";" Or Left$(ln, 1) = "'" Then
                ' blank line or comment
            ElseIf Left$(ln, 1) = "[" Then
                section = UCase$(Trim$(Replace(Replace(ln, "[", ""), "]", "")))
            ElseIf section = "MODEL" Or section = "DRAWING" Then
                ParseIniLine section, ln
            End If
        Loop
        ts.Close
    End If
Fallback:
    If gFormatCount = 0 Then LoadDefaultFormats
End Sub

Private Sub ParseIniLine(ByVal section As String, ByVal ln As String)
    Dim parts() As String, k As Long, flg As String, onOff As String
    parts = Split(ln, ",")
    If UBound(parts) < 1 Then Exit Sub
    For k = 0 To UBound(parts)
        parts(k) = Trim$(Replace(parts(k), """", ""))
    Next k
    If parts(1) = "" Then Exit Sub
    If UBound(parts) >= 2 Then flg = UCase$(parts(2))
    If UBound(parts) >= 3 Then onOff = UCase$(parts(3))
    If flg = "" Then flg = IIf(section = "MODEL", "MESH", "MULTI")
    AddFormat parts(0), LCase$(parts(1)), section, flg, (onOff = "ON")
End Sub

Private Sub AddFormat(ByVal pLabel As String, ByVal pExt As String, ByVal pKind As String, _
                      ByVal pFlag As String, ByVal pOn As Boolean)
    If Left$(pExt, 1) <> "." Then pExt = "." & pExt
    If gFormatCount > UBound(gFormats) Then ReDim Preserve gFormats(0 To gFormatCount + 16)
    With gFormats(gFormatCount)
        .Label = pLabel
        .Ext = pExt
        .Kind = pKind
        .Flag = pFlag
        .DefaultOn = pOn
        .Selected = pOn
    End With
    gFormatCount = gFormatCount + 1
End Sub

' Built-in defaults when SWExporter.ini is missing or unreadable: only STEP is pre-ticked.
Private Sub LoadDefaultFormats()
    AddFormat "STL", ".stl", "MODEL", "MESH", False
    AddFormat "STEP AP214", ".step", "MODEL", "ISOLATE", True
    AddFormat "IGES", ".igs", "MODEL", "ISOLATE", False
    AddFormat "Parasolid", ".x_t", "MODEL", "ISOLATE", False
    AddFormat "ACIS SAT", ".sat", "MODEL", "ISOLATE", False
    AddFormat "3MF", ".3mf", "MODEL", "MESH", False
    AddFormat "GLB", ".glb", "MODEL", "MESH", False
    AddFormat "OBJ", ".obj", "MODEL", "MESH", False
    AddFormat "JPEG image", ".jpg", "MODEL", "MESH", False
    AddFormat "PNG image", ".png", "MODEL", "MESH", False
    AddFormat "eDrawings", ".e*", "MODEL", "WHOLE", False
    AddFormat "3D PDF", ".pdf", "MODEL", "WHOLE", False
    AddFormat "PDF", ".pdf", "DRAWING", "SINGLE", False
    AddFormat "eDrawings", ".edrw", "DRAWING", "SINGLE", False
    AddFormat "DWG", ".dwg", "DRAWING", "MULTI", False
    AddFormat "DXF", ".dxf", "DRAWING", "MULTI", False
    AddFormat "JPEG", ".jpg", "DRAWING", "MULTI", False
    AddFormat "PNG", ".png", "DRAWING", "MULTI", False
End Sub

'============================ LIST HELPERS (used by the form) =================
Public Function GetBodyNames(m As SldWorks.ModelDoc2, ByVal inclSurf As Boolean) As Collection
    Dim c As Collection, v As Variant, i As Long
    Set c = New Collection
    v = VisibleBodies(m, inclSurf)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            c.Add CStr(v(i).Name)
        Next i
    End If
    Set GetBodyNames = c
End Function

' Variant array of visible Body2 objects (solids first, then surfaces if asked), or Empty.
Public Function VisibleBodies(m As SldWorks.ModelDoc2, ByVal inclSurf As Boolean) As Variant
    Dim p As SldWorks.PartDoc, vS As Variant, vF As Variant
    Dim out() As Object, n As Long, i As Long
    If m Is Nothing Then Exit Function
    If m.GetType <> swDocumentTypes_e.swDocPART Then Exit Function
    Set p = m
    vS = p.GetBodies2(swBodyType_e.swSolidBody, True)
    If inclSurf Then vF = p.GetBodies2(swBodyType_e.swSheetBody, True)
    If IsArray(vS) Then
        For i = LBound(vS) To UBound(vS)
            ReDim Preserve out(0 To n)
            Set out(n) = vS(i)
            n = n + 1
        Next i
    End If
    If IsArray(vF) Then
        For i = LBound(vF) To UBound(vF)
            ReDim Preserve out(0 To n)
            Set out(n) = vF(i)
            n = n + 1
        Next i
    End If
    If n > 0 Then VisibleBodies = out
End Function

Public Function MatchesFilter(ByVal s As String, ByVal f As String, ByVal caseSens As Boolean) As Boolean
    If f = "" Then
        MatchesFilter = True
    ElseIf caseSens Then
        MatchesFilter = (InStr(1, s, f, vbBinaryCompare) > 0)
    Else
        MatchesFilter = (InStr(1, s, f, vbTextCompare) > 0)
    End If
End Function

Public Function ListModelFiles(ByVal folder As String, ByVal recurse As Boolean) As Collection
    Dim c As Collection
    Set c = New Collection
    On Error Resume Next
    If folder <> "" Then
        If FS().FolderExists(folder) Then AddFiles FS().GetFolder(folder), recurse, c
    End If
    Set ListModelFiles = c
End Function

Private Sub AddFiles(ByVal fld As Object, ByVal recurse As Boolean, ByVal c As Collection)
    Dim f As Object, sf As Object, e As String
    On Error Resume Next          ' skip folders we are not allowed to read
    For Each f In fld.Files
        e = LCase$(Mid$(f.Name, InStrRev(f.Name, ".") + 1))
        If (e = "sldprt" Or e = "sldasm" Or e = "slddrw") And Left$(f.Name, 2) <> "~$" Then c.Add CStr(f.Path)
    Next f
    If recurse Then
        For Each sf In fld.SubFolders
            AddFiles sf, True, c
        Next sf
    End If
End Sub

' Multi-select Open dialog. Returns full paths (empty Collection if cancelled).
Public Function PickFiles(ByVal title As String, ByVal initDir As String) As Collection
    Dim ofn As OPENFILENAME, s As String, parts() As String, baseDir As String
    Dim i As Long, c As Collection
    Set c = New Collection
    Set PickFiles = c
    ofn.lStructSize = LenB(ofn)
    ofn.hwndOwner = GetActiveWindow()
    ofn.lpstrFilter = "SOLIDWORKS files (*.sldprt;*.sldasm;*.slddrw)" & vbNullChar & "*.sldprt;*.sldasm;*.slddrw" & vbNullChar & _
                      "Parts (*.sldprt)" & vbNullChar & "*.sldprt" & vbNullChar & _
                      "Assemblies (*.sldasm)" & vbNullChar & "*.sldasm" & vbNullChar & _
                      "Drawings (*.slddrw)" & vbNullChar & "*.slddrw" & vbNullChar & vbNullChar
    ofn.nFilterIndex = 1
    ofn.lpstrFile = String$(65536, vbNullChar)
    ofn.nMaxFile = 65536
    If initDir <> "" Then ofn.lpstrInitialDir = initDir
    ofn.lpstrTitle = title
    ' OFN_EXPLORER | OFN_ALLOWMULTISELECT | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY
    ofn.flags = &H80000 Or &H200 Or &H1000 Or &H800 Or &H4
    If GetOpenFileNameA(ofn) = 0 Then Exit Function
    s = ofn.lpstrFile
    i = InStr(s, vbNullChar & vbNullChar)
    If i > 0 Then s = Left$(s, i - 1)
    s = Replace(s, vbNullChar & vbNullChar, "")
    If s = "" Then Exit Function
    parts = Split(s, vbNullChar)
    If UBound(parts) = 0 Then
        c.Add parts(0)                              ' one file: full path
    Else
        baseDir = parts(0)                          ' several: folder, then names
        If Right$(baseDir, 1) <> "\" Then baseDir = baseDir & "\"
        For i = 1 To UBound(parts)
            If parts(i) <> "" Then c.Add baseDir & parts(i)
        Next i
    End If
End Function

'============================ ASSEMBLY HELPERS ================================
' Unique parts of an assembly (all levels), keyed "path|referenced config",
' value = first Component2 found. Suppressed and envelope components are left out;
' visibleOnly also leaves out hidden components (or ones inside a hidden subassembly).
Public Function AssemblyParts(asm As SldWorks.ModelDoc2, ByVal visibleOnly As Boolean) As Object
    Dim d As Object, a As Object, v As Variant, i As Long, comp As Object, key As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set AssemblyParts = d
    If asm Is Nothing Then Exit Function
    If asm.GetType <> swDocumentTypes_e.swDocASSEMBLY Then Exit Function
    Set a = asm
    v = a.GetComponents(False)
    If Not IsArray(v) Then Exit Function
    For i = LBound(v) To UBound(v)
        Set comp = v(i)
        If UsablePartComp(comp, visibleOnly) Then
            key = comp.GetPathName & "|" & comp.ReferencedConfiguration
            If Not d.Exists(key) Then d.Add key, comp
        End If
    Next i
End Function

Public Function PathOfKey(ByVal key As String) As String
    If InStr(key, "|") > 0 Then PathOfKey = Left$(key, InStr(key, "|") - 1) Else PathOfKey = key
End Function

Public Function CfgOfKey(ByVal key As String) As String
    If InStr(key, "|") > 0 Then CfgOfKey = Mid$(key, InStr(key, "|") + 1)
End Function

Public Function PartKeyDisplay(ByVal key As String) As String
    Dim nm As String
    nm = PathOfKey(key)
    nm = Mid$(nm, InStrRev(nm, "\") + 1)
    If InStrRev(nm, ".") > 0 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
    PartKeyDisplay = nm & "  [" & CfgOfKey(key) & "]"
End Function

Private Function UsablePartComp(ByVal comp As Object, ByVal visibleOnly As Boolean) As Boolean
    Dim p As String, b As Boolean
    On Error GoTo Nope
    If comp Is Nothing Then Exit Function
    p = comp.GetPathName
    If LCase$(Right$(p, 7)) <> ".sldprt" Then Exit Function
    If comp.IsSuppressed Then Exit Function
    b = False
    On Error Resume Next
    b = comp.IsEnvelope
    On Error GoTo Nope
    If b Then Exit Function
    If visibleOnly Then
        If Not IsCompShown(comp) Then Exit Function
    End If
    UsablePartComp = True
Nope:
End Function

Private Function IsCompShown(ByVal comp As Object) As Boolean
    Dim c As Object, p As Object, h As Boolean, guard As Long
    On Error Resume Next
    Set c = comp
    Do While Not c Is Nothing
        h = False
        h = c.IsHidden(True)
        If h Then Exit Function
        Set p = Nothing
        Set p = c.GetParent
        Set c = p
        guard = guard + 1
        If guard > 64 Then Exit Do
    Loop
    IsCompShown = True
End Function

'------------------------- visible bodies in assembly context -------------------
' Walks the assembly tree from the root component of the active configuration.
' Skips (with everything below them) components that are suppressed, hidden in the
' active display state, or envelopes. ExcludeFromBOM is NOT a visibility setting
' and is ignored. For every remaining part component it reads the bodies of that
' component instance (Component2.GetBodies3) and keeps the ones whose Body2.Visible
' is True in the assembly context, so bodies hidden in the assembly, by the active
' display state, or in the part file itself are all left out.
' Returns a Dictionary "path|referenced config" -> Collection of body names (union
' over every shown instance of that part + configuration). comps receives a
' Dictionary with the same keys -> first Component2 found.
Public Function VisibleBodyMap(asm As SldWorks.ModelDoc2, ByVal inclSurf As Boolean, comps As Object) As Object
    Dim map As Object, root As Object, v As Variant, a As SldWorks.AssemblyDoc
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    Set comps = CreateObject("Scripting.Dictionary")
    comps.CompareMode = vbTextCompare
    Set VisibleBodyMap = map
    If asm Is Nothing Then Exit Function
    If asm.GetType <> swDocumentTypes_e.swDocASSEMBLY Then Exit Function
    On Error Resume Next
    Set root = asm.ConfigurationManager.ActiveConfiguration.GetRootComponent3(False)
    If Not root Is Nothing Then
        v = Empty
        v = root.GetChildren
    Else
        Set a = asm
        v = Empty
        v = a.GetComponents(True)                   ' top level only; children are walked below
    End If
    On Error GoTo 0
    CollectVisibleBodies v, inclSurf, map, comps, 0
End Function

Private Sub CollectVisibleBodies(ByVal v As Variant, ByVal inclSurf As Boolean, ByVal map As Object, _
                                 ByVal comps As Object, ByVal depth As Long)
    Dim i As Long, ch As Object, p As String, kids As Variant
    If Not IsArray(v) Then Exit Sub
    If depth > 64 Then Exit Sub
    For i = LBound(v) To UBound(v)
        Set ch = Nothing
        On Error Resume Next
        Set ch = v(i)
        On Error GoTo 0
        If CompIsShownHere(ch) Then
            p = ""
            On Error Resume Next
            p = LCase$(ch.GetPathName)
            On Error GoTo 0
            If Right$(p, 7) = ".sldprt" Then
                AddCompVisibleBodies ch, inclSurf, map, comps
            ElseIf Right$(p, 7) = ".sldasm" Then
                kids = Empty
                On Error Resume Next
                kids = ch.GetChildren
                On Error GoTo 0
                CollectVisibleBodies kids, inclSurf, map, comps, depth + 1
            End If
        End If
    Next i
End Sub

' This component itself (its parents were already checked by the tree walk).
Private Function CompIsShownHere(ByVal comp As Object) As Boolean
    Dim b As Boolean, vis As Long
    On Error GoTo Nope
    If comp Is Nothing Then Exit Function
    If comp.IsSuppressed Then Exit Function
    vis = 1
    vis = comp.Visible                  ' swComponentVisibilityState_e: 0 = hidden (active display state)
    If vis = 0 Then Exit Function
    If comp.IsHidden(True) Then Exit Function
    b = False
    On Error Resume Next
    b = comp.IsEnvelope
    On Error GoTo Nope
    If b Then Exit Function
    CompIsShownHere = True
Nope:
End Function

Private Sub AddCompVisibleBodies(ByVal comp As Object, ByVal inclSurf As Boolean, ByVal map As Object, ByVal comps As Object)
    Dim key As String, c As Collection, vb As Variant, info As Variant, j As Long
    Dim t As Long, nT As Long, st As Long, vis As Boolean, b As Object, nm As String
    On Error Resume Next
    key = comp.GetPathName & "|" & comp.ReferencedConfiguration
    If map.Exists(key) Then
        Set c = map(key)
    Else
        Set c = New Collection
        map.Add key, c
        comps.Add key, comp
    End If
    If inclSurf Then nT = 1 Else nT = 0
    For t = 0 To nT                     ' 0 = swSolidBody, 1 = swSheetBody
        vb = Empty
        vb = comp.GetBodies3(t, info)
        If (Not IsArray(vb)) And t = 0 Then
            st = -1
            st = comp.GetSuppression2
            If st = 1 Or st = 4 Then    ' lightweight: resolve it so its bodies can be read
                comp.SetSuppression2 2  ' swComponentFullyResolved
                vb = Empty
                vb = comp.GetBodies3(t, info)
            End If
        End If
        If IsArray(vb) Then
            For j = LBound(vb) To UBound(vb)
                Set b = Nothing
                Set b = vb(j)
                If Not b Is Nothing Then
                    vis = False
                    vis = b.Visible     ' Body2.Visible of the component instance = what the assembly shows
                    If vis Then
                        nm = BodyKeyName(CStr(b.Name))
                        If nm <> "" Then AddUnique c, nm
                    End If
                End If
            Next j
        End If
    Next t
End Sub

' Body name as the part document knows it (drops any "comp/" prefix or "@..." suffix).
Public Function BodyKeyName(ByVal n As String) As String
    If InStrRev(n, "/") > 0 Then n = Mid$(n, InStrRev(n, "/") + 1)
    If InStr(n, "@") > 0 Then n = Left$(n, InStr(n, "@") - 1)
    BodyKeyName = n
End Function

Public Sub AddUnique(ByVal c As Collection, ByVal s As String)
    If Not InCollection(c, s) Then c.Add s
End Sub

Private Function CompShortName(ByVal comp As Object) As String
    Dim n As String
    On Error Resume Next
    n = comp.Name2
    If InStrRev(n, "/") > 0 Then n = Mid$(n, InStrRev(n, "/") + 1)
    CompShortName = n
End Function

' Gets the component's part document and makes sure it has a window.
' closeAfter: 0 = leave as is, 1 = CloseDoc afterwards, 2 = set Visible = False afterwards.
Private Function OpenCompPart(ByVal comp As Object, closeAfter As Long) As SldWorks.ModelDoc2
    Dim p As String, pm As SldWorks.ModelDoc2, pm2 As SldWorks.ModelDoc2, e As Long, w As Long, vis As Boolean
    closeAfter = 0
    On Error Resume Next
    p = comp.GetPathName
    Set pm = swApp.GetOpenDocumentByName(p)
    If Not pm Is Nothing Then
        vis = pm.Visible
        If vis Then
            Set OpenCompPart = pm
            Exit Function
        End If
    End If
    If InStr(p, "^") = 0 Then                       ' not a virtual component
        Set pm2 = swApp.OpenDoc6(p, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", e, w)
        If Not pm2 Is Nothing Then
            closeAfter = 1
            Set OpenCompPart = pm2
            Exit Function
        End If
    End If
    If pm Is Nothing Then Set pm = comp.GetModelDoc2
    If Not pm Is Nothing Then
        pm.Visible = True
        closeAfter = 2
    End If
    Set OpenCompPart = pm
End Function

'================================ RUN ==========================================
Public Sub RunExport(m As SldWorks.ModelDoc2)
    ' m = active document (Nothing in batch mode with no document open)
    Dim oldBin As Boolean, oldOneFile As Boolean, oldUnits As Long, oldQual As Long
    Dim oldDxf As Long, oldStep As Long, origActive As SldWorks.ModelDoc2
    Dim t0 As Date, logDir As String, logFile As String, msg As String, e As Long

    Set mPlanned = New Collection
    Set mDone = CreateObject("Scripting.Dictionary")
    mDone.CompareMode = vbTextCompare
    Set mAnchorDoc = Nothing
    mCompName = ""
    Set mBodyFilter = Nothing
    mFilterStrict = False
    mLog = "time,source,config,body,sheet,format,output,status,errors,warnings" & vbCrLf
    mOk = 0: mFail = 0: mSkip = 0: mStep = 0
    gCancel = False
    mFirstOutDir = "": mFirstSrcDir = ""
    mRunStamp = Format$(Now, "yyyymmdd_hhnnss")
    t0 = Now
    Set origActive = swApp.ActiveDoc

    ' remember the user's export options, then set ours for this run
    oldBin = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swSTLBinaryFormat)
    oldOneFile = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swSTLComponentsIntoOneFile)
    oldUnits = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swExportStlUnits)
    oldQual = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSTLQuality)
    oldDxf = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swDxfMultiSheetOption)
    oldStep = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP)

    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSTLBinaryFormat, gStlBinary
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSTLComponentsIntoOneFile, True   ' assemblies: one STL, not one per component
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swExportStlUnits, gStlUnits
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swSTLQuality, gStlQuality
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, 0  ' swDxfActiveSheetOnly
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swStepAP, 214            ' the label says AP214

    On Error GoTo Fail
    If gMode = MODE_BATCH Then
        RunBatch
    ElseIf m.GetType = swDocumentTypes_e.swDocDRAWING Then
        ExportDrawing m, gSelItems
    ElseIf gMode = MODE_COMPONENTS Or gMode = MODE_COMPBODIES Then
        ExportComponents m, gMode, gSelItems
    ElseIf gMode = MODE_VISBODIES Then
        If m.GetType = swDocumentTypes_e.swDocASSEMBLY Then
            ExportVisibleBodies m, gSelItems
        Else
            ExportModel m, MODE_BODIES, gSelItems       ' part: the ticked visible bodies, like Bodies mode
        End If
    Else
        ExportModel m, gMode, gSelItems
    End If
    GoTo Done
Fail:
    AddLog "", "", "", "", "", "", "STOPPED: " & Err.Description, 0, 0
    mFail = mFail + 1
    Resume Done
Done:
    On Error Resume Next
    Set mAnchorDoc = Nothing
    mCompName = ""
    Set mBodyFilter = Nothing
    mFilterStrict = False
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSTLBinaryFormat, oldBin
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSTLComponentsIntoOneFile, oldOneFile
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swExportStlUnits, oldUnits
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swSTLQuality, oldQual
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, oldDxf
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swStepAP, oldStep
    If Not origActive Is Nothing Then
        swApp.ActivateDoc3 DocKey(origActive), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, e
    End If

    If gWriteLog Then
        logDir = ""
        If gMode = MODE_BATCH Then logDir = gBatchFolder
        If logDir <> "" Then
            If Not FS().FolderExists(logDir) Then logDir = ""
        End If
        If logDir = "" Then logDir = mFirstOutDir
        If logDir <> "" Then
            If Not FS().FolderExists(logDir) Then logDir = mFirstSrcDir   ' e.g. dry run into a folder that doesn't exist yet
        End If
        If logDir = "" Then logDir = mFirstSrcDir
        If logDir = "" Then logDir = MacroFolder()
        If logDir <> "" Then
            logFile = logDir & "\SWExporter_log_" & mRunStamp & IIf(gDryRun, "_DRYRUN", "") & ".csv"
            If Not WriteText(logFile, mLog) Then logFile = ""
        End If
    End If
    SetStatus ""

    msg = IIf(gDryRun, "DRY RUN - nothing was written." & vbCrLf & vbCrLf, "") & _
          mOk & IIf(gDryRun, " file(s) would be written", " file(s) written") & ", " & _
          mSkip & " skipped, " & mFail & " failed in " & _
          Format$(DateDiff("s", t0, Now) / 60, "0.00") & " min." & _
          IIf(gCancel, vbCrLf & "Cancelled by user.", "")
    If logFile <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Log: " & logFile
    ElseIf gWriteLog Then
        msg = msg & vbCrLf & vbCrLf & "The log file could not be written."
    End If
    MsgBox msg, IIf(mFail = 0, vbInformation, vbExclamation), "SWExporter"
End Sub

'---------------------------- batch --------------------------------------------
Private Sub RunBatch()
    Dim i As Long
    For i = 1 To gSelItems.Count
        If gCancel Then Exit For
        SetStatus "Batch " & i & "/" & gSelItems.Count & ": " & gSelItems(i)
        BatchOne CStr(gSelItems(i))
        DoEvents
    Next i
End Sub

' One file. Its own error handler, so a bad file is logged and the batch goes on.
Private Sub BatchOne(ByVal path As String)
    Dim m As SldWorks.ModelDoc2, drw As SldWorks.DrawingDoc
    Dim wasOpen As Boolean, t As Long, e As Long, w As Long, bm As Long, k As Long
    Dim v As Variant, items As Collection, doneKey As String

    Select Case LCase$(Mid$(path, InStrRev(path, ".") + 1))
        Case "sldprt": t = swDocumentTypes_e.swDocPART
        Case "sldasm": t = swDocumentTypes_e.swDocASSEMBLY
        Case "slddrw": t = swDocumentTypes_e.swDocDRAWING
        Case Else: Exit Sub
    End Select

    On Error GoTo EH
    Set m = swApp.GetOpenDocumentByName(path)
    wasOpen = Not (m Is Nothing)
    If wasOpen Then wasOpen = m.Visible        ' loaded only as a reference = treat as not open
    If Not wasOpen Then
        Set m = swApp.OpenDoc6(path, t, swOpenDocOptions_e.swOpenDocOptions_Silent, "", e, w)
    End If
    If m Is Nothing Then
        AddLog path, "", "", "", "", "", "OPEN FAILED", e, w
        mFail = mFail + 1
        Exit Sub
    End If
    swApp.ActivateDoc3 DocKey(m), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, e

    Set items = New Collection
    If t = swDocumentTypes_e.swDocDRAWING Then
        Set drw = m
        v = drw.GetSheetNames
        If IsArray(v) Then
            For k = LBound(v) To UBound(v)
                items.Add CStr(v(k))
            Next k
        End If
        ExportDrawing m, items
    ElseIf t = swDocumentTypes_e.swDocASSEMBLY Then
        ExportAssemblyParts m                     ' every part inside, split per "Parts in batch"
    Else
        bm = gBatchModelMode
        Select Case bm
            Case MODE_BODIES: doneKey = path & "|" & m.ConfigurationManager.ActiveConfiguration.Name & "|B"
            Case MODE_CFGBODIES: doneKey = path & "|*|CB"
            Case Else: doneKey = path & "|*|C"
        End Select
        If mDone.Exists(doneKey) Then
            AddLog path, "", "", "", "", "", "SKIPPED: already exported in this run", 0, 0
            mSkip = mSkip + 1
        Else
            mDone.Add doneKey, True
            If bm = MODE_BODIES Then
                ExportModel m, MODE_BODIES, Nothing
            Else
                v = m.GetConfigurationNames
                If IsArray(v) Then
                    For k = LBound(v) To UBound(v)
                        If MatchesFilter(CStr(v(k)), gCfgFilter, gCfgFilterCase) Then items.Add CStr(v(k))
                    Next k
                End If
                If items.Count = 0 Then
                    AddLog path, "", "", "", "", "", "SKIPPED: no configuration matches the filter", 0, 0
                    mSkip = mSkip + 1
                Else
                    ExportModel m, bm, items
                End If
            End If
        End If
    End If
Cleanup:
    On Error Resume Next
    If (Not wasOpen) And (Not m Is Nothing) Then swApp.CloseDoc DocKey(m)    ' closed without saving
    Exit Sub
EH:
    AddLog path, "", "", "", "", "", "ERROR: " & Err.Description, 0, 0
    mFail = mFail + 1
    Resume Cleanup
End Sub

' Batch: every part in the assembly (all levels), once per run, split per "Parts in batch".
Private Sub ExportAssemblyParts(asm As SldWorks.ModelDoc2)
    Dim d As Object, k As Variant, p As String, cfg As String, doneKey As String
    Dim spec As String, md As Long, n As Long
    Set d = AssemblyParts(asm, False)
    If d.Count = 0 Then
        AddLog DocKey(asm), "", "", "", "", "", "SKIPPED: no parts found in the assembly", 0, 0
        mSkip = mSkip + 1
        Exit Sub
    End If
    For Each k In d.Keys
        If gCancel Then Exit For
        p = PathOfKey(CStr(k))
        cfg = CfgOfKey(CStr(k))
        Select Case gBatchModelMode
            Case MODE_BODIES                       ' bodies of the configuration the assembly uses
                doneKey = p & "|" & cfg & "|B": spec = cfg: md = MODE_CFGBODIES
            Case MODE_CFGBODIES
                doneKey = p & "|*|CB": spec = "*": md = MODE_CFGBODIES
            Case Else
                doneKey = p & "|*|C": spec = "*": md = MODE_CONFIGS
        End Select
        If Not mDone.Exists(doneKey) Then
            mDone.Add doneKey, True
            n = n + 1
            SetStatus "Assembly part " & n & ": " & PartKeyDisplay(CStr(k))
            ExportAsmPart asm, d(k), spec, md
        End If
        DoEvents
    Next k
End Sub

' Components / Components x bodies with an assembly open.
' Bodies hidden in the assembly (or its display state) are left out of each part.
Private Sub ExportComponents(asm As SldWorks.ModelDoc2, ByVal mode As Long, ByVal sel As Collection)
    Dim d As Object, k As Variant, n As Long, md As Long
    Dim vmap As Object, vcomps As Object, filt As Collection
    If mode = MODE_COMPONENTS Then md = MODE_CONFIGS Else md = MODE_CFGBODIES
    Set d = AssemblyParts(asm, True)
    Set vmap = VisibleBodyMap(asm, True, vcomps)        ' solids + surfaces shown in the assembly
    For Each k In d.Keys
        If gCancel Then Exit For
        If InCollection(sel, CStr(k)) Then
            n = n + 1
            SetStatus "Component " & n & ": " & PartKeyDisplay(CStr(k))
            Set filt = Nothing
            If vmap.Exists(CStr(k)) Then
                Set filt = vmap(CStr(k))
                If filt.Count = 0 Then Set filt = Nothing   ' bodies could not be read: old behaviour
            End If
            mFilterStrict = False
            ExportAsmPart asm, d(k), CfgOfKey(CStr(k)), md, filt
        End If
        DoEvents
    Next k
    If n = 0 Then
        AddLog DocKey(asm), "", "", "", "", "", "SKIPPED: none of the ticked parts were found", 0, 0
        mSkip = mSkip + 1
    End If
End Sub

' Visible bodies with an assembly open. sel holds "path|config" & vbTab & body name.
Private Sub ExportVisibleBodies(asm As SldWorks.ModelDoc2, ByVal sel As Collection)
    Dim map As Object, comps As Object, k As Variant, filt As Collection
    Dim i As Long, n As Long, s As String, pfx As String
    Set map = VisibleBodyMap(asm, gIncludeSurfaces, comps)
    For Each k In map.Keys
        If gCancel Then Exit For
        pfx = CStr(k) & vbTab
        Set filt = New Collection
        For i = 1 To sel.Count
            s = CStr(sel(i))
            If Len(s) > Len(pfx) Then
                If StrComp(Left$(s, Len(pfx)), pfx, vbTextCompare) = 0 Then AddUnique filt, Mid$(s, Len(pfx) + 1)
            End If
        Next i
        If filt.Count > 0 Then
            n = n + 1
            SetStatus "Visible bodies " & n & ": " & PartKeyDisplay(CStr(k))
            mFilterStrict = True
            ExportAsmPart asm, comps(k), CfgOfKey(CStr(k)), MODE_CFGBODIES, filt
        End If
        DoEvents
    Next k
    mFilterStrict = False
    If n = 0 Then
        AddLog DocKey(asm), "", "", "", "", "", "SKIPPED: none of the ticked bodies were found", 0, 0
        mSkip = mSkip + 1
    End If
End Sub

' Exports one part of an assembly from its own part document.
' bodyFilter = names of the bodies shown in the assembly (Nothing = every visible body of the part).
' cfgSpec = a configuration name, or "*" for all configurations matching the config filter.
' Output folder and {file} in the subfolder name follow the assembly.
Private Sub ExportAsmPart(asm As SldWorks.ModelDoc2, ByVal comp As Object, ByVal cfgSpec As String, ByVal mode As Long, _
                          Optional ByVal bodyFilter As Collection = Nothing)
    Dim pm As SldWorks.ModelDoc2, closeAfter As Long, cfgs As Collection, v As Variant
    Dim k As Long, e As Long, p As String
    On Error GoTo EH
    p = comp.GetPathName
    Set pm = OpenCompPart(comp, closeAfter)
    If pm Is Nothing Then
        AddLog p, cfgSpec, "", "", "", "", "OPEN FAILED", 0, 0
        mFail = mFail + 1
        GoTo Cleanup
    End If
    swApp.ActivateDoc3 DocKey(pm), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, e

    Set cfgs = New Collection
    If cfgSpec = "*" Then
        v = pm.GetConfigurationNames
        If IsArray(v) Then
            For k = LBound(v) To UBound(v)
                If MatchesFilter(CStr(v(k)), gCfgFilter, gCfgFilterCase) Then cfgs.Add CStr(v(k))
            Next k
        End If
    Else
        cfgs.Add cfgSpec
    End If
    If cfgs.Count = 0 Then
        AddLog DocKey(pm), "", "", "", "", "", "SKIPPED: no configuration matches the filter", 0, 0
        mSkip = mSkip + 1
    Else
        Set mAnchorDoc = asm
        mCompName = CompShortName(comp)
        Set mBodyFilter = bodyFilter
        ExportModel pm, mode, cfgs
    End If
Cleanup:
    On Error Resume Next
    Set mAnchorDoc = Nothing
    mCompName = ""
    Set mBodyFilter = Nothing
    If Not pm Is Nothing Then
        If closeAfter = 1 Then swApp.CloseDoc DocKey(pm)
        If closeAfter = 2 Then pm.Visible = False
    End If
    swApp.ActivateDoc3 DocKey(asm), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, e
    Exit Sub
EH:
    AddLog p, cfgSpec, "", "", "", "", "ERROR: " & Err.Description, 0, 0
    mFail = mFail + 1
    Resume Cleanup
End Sub

'---------------------------- parts / assemblies --------------------------------
Private Sub ExportModel(m As SldWorks.ModelDoc2, ByVal mode As Long, ByVal sel As Collection)
    Dim origCfg As String, selObjs As Collection, cfgList As Collection, bodySel As Collection
    Dim i As Long, f As Long, cfgName As String

    origCfg = m.ConfigurationManager.ActiveConfiguration.Name
    Set selObjs = SaveSelection(m)
    mModelMode = mode

    Set cfgList = New Collection
    If mode = MODE_BODIES Or (sel Is Nothing) Then
        cfgList.Add origCfg
    Else
        For i = 1 To sel.Count
            cfgList.Add CStr(sel(i))
        Next i
    End If
    If mode = MODE_BODIES Then Set bodySel = sel       ' Nothing = all bodies

    On Error GoTo EH
    For i = 1 To cfgList.Count
        If gCancel Then Exit For
        cfgName = cfgList(i)
        If cfgName <> m.ConfigurationManager.ActiveConfiguration.Name Then
            If Not m.ShowConfiguration2(cfgName) Then
                AddLog DocKey(m), cfgName, "", "", "", "", "FAILED: could not activate configuration", 0, 0
                mFail = mFail + 1
                GoTo NextCfg
            End If
            m.GraphicsRedraw2
        End If
        If mode = MODE_CONFIGS Or m.GetType <> swDocumentTypes_e.swDocPART Then
            ' whole model per configuration
            If mode = MODE_CONFIGS And (Not mBodyFilter Is Nothing) And m.GetType = swDocumentTypes_e.swDocPART Then
                ExportWholeFiltered m, cfgName              ' assembly component: leave out bodies hidden in the assembly
            Else
                For f = 0 To gFormatCount - 1
                    If gFormats(f).Selected And gFormats(f).Kind = "MODEL" Then
                        SaveOne m, gFormats(f), cfgName, "", ""
                    End If
                Next f
            End If
        Else
            ExportBodies m, cfgName, bodySel
        End If
NextCfg:
        DoEvents
    Next i
Cleanup:
    On Error Resume Next
    If m.ConfigurationManager.ActiveConfiguration.Name <> origCfg Then m.ShowConfiguration2 origCfg
    RestoreSelection m, selObjs
    m.GraphicsRedraw2
    Exit Sub
EH:
    AddLog DocKey(m), cfgName, "", "", "", "", "ERROR: " & Err.Description, 0, 0
    mFail = mFail + 1
    Resume Cleanup
End Sub

Private Sub ExportBodies(m As SldWorks.ModelDoc2, ByVal cfgName As String, ByVal bodySel As Collection)
    Dim vAll As Variant, b As SldWorks.Body2, i As Long, j As Long, f As Long
    Dim bName As String, needMesh As Boolean, nFound As Long, filt As Collection

    ' vAll = every visible body (solids + surfaces). All of them are hidden while one
    ' body is exported to a mesh format, so nothing else can leak into the file.
    ' filt (assembly export) = bodies shown in the assembly; the others are skipped.
    vAll = VisibleBodies(m, True)
    Set filt = mBodyFilter
    If IsArray(vAll) Then
        For i = LBound(vAll) To UBound(vAll)
            Set b = vAll(i)
            If gIncludeSurfaces Or b.GetType = swBodyType_e.swSolidBody Then
                If InCollection(filt, CStr(b.Name)) Then nFound = nFound + 1
            End If
        Next i
        If nFound = 0 And (Not filt Is Nothing) And (Not mFilterStrict) Then
            ' no body name matched the assembly's list: fall back to every visible body
            Set filt = Nothing
            For i = LBound(vAll) To UBound(vAll)
                Set b = vAll(i)
                If gIncludeSurfaces Or b.GetType = swBodyType_e.swSolidBody Then nFound = nFound + 1
            Next i
        End If
    End If
    If nFound = 0 Then
        AddLog DocKey(m), cfgName, "", "", "", "", "SKIPPED: no visible bodies", 0, 0
        mSkip = mSkip + 1
        Exit Sub
    End If
    For f = 0 To gFormatCount - 1
        If gFormats(f).Selected And gFormats(f).Kind = "MODEL" And gFormats(f).Flag = "MESH" Then needMesh = True
    Next f

    On Error GoTo EH
    For i = LBound(vAll) To UBound(vAll)
        If gCancel Then Exit For
        Set b = vAll(i)
        bName = b.Name
        If (gIncludeSurfaces Or b.GetType = swBodyType_e.swSolidBody) And InCollection(bodySel, bName) _
           And InCollection(filt, bName) Then
            ' 1) mesh / image formats: hide every other body, save, show again
            If needMesh Then
                For j = LBound(vAll) To UBound(vAll)
                    vAll(j).HideBody (j <> i)
                Next j
                m.GraphicsRedraw2
                For f = 0 To gFormatCount - 1
                    If gFormats(f).Selected And gFormats(f).Kind = "MODEL" And gFormats(f).Flag = "MESH" Then
                        SaveOne m, gFormats(f), cfgName, bName, ""
                    End If
                Next f
                ShowAll vAll
            End If
            ' 2) B-rep formats: copy the body into one temporary part, save each format
            SaveIsolated m, b, cfgName, bName
        End If
    Next i

    ' 3) WHOLE formats (eDrawings / 3D PDF) once per configuration
    If Not gCancel Then
        For f = 0 To gFormatCount - 1
            If gFormats(f).Selected And gFormats(f).Kind = "MODEL" And gFormats(f).Flag = "WHOLE" Then
                SaveOne m, gFormats(f), cfgName, "", ""
            End If
        Next f
    End If
Cleanup:
    On Error Resume Next
    ShowAll vAll
    m.GraphicsRedraw2
    Exit Sub
EH:
    AddLog DocKey(m), cfgName, bName, "", "", "", "ERROR: " & Err.Description, 0, 0
    mFail = mFail + 1
    Resume Cleanup
End Sub

Private Sub ShowAll(ByVal v As Variant)
    Dim j As Long
    On Error Resume Next
    If Not IsArray(v) Then Exit Sub
    For j = LBound(v) To UBound(v)
        v(j).HideBody False
    Next j
End Sub

' Whole part, minus the bodies hidden in the assembly (mBodyFilter = bodies shown there).
' Mesh / image / eDrawings / 3D PDF: the hidden bodies are hidden in the part while it saves.
' STEP / IGES / Parasolid / SAT: the shown bodies are copied into one temporary part.
' When nothing is hidden in the assembly, this is the normal whole-model export.
Private Sub ExportWholeFiltered(m As SldWorks.ModelDoc2, ByVal cfgName As String)
    Dim vAll As Variant, b As SldWorks.Body2, i As Long, f As Long
    Dim keep() As Object, hid() As Object, nKeep As Long, nHide As Long, vKeep As Variant
    vAll = VisibleBodies(m, True)
    If IsArray(vAll) Then
        For i = LBound(vAll) To UBound(vAll)
            Set b = vAll(i)
            If InCollection(mBodyFilter, CStr(b.Name)) Then
                ReDim Preserve keep(0 To nKeep)
                Set keep(nKeep) = b
                nKeep = nKeep + 1
            Else
                ReDim Preserve hid(0 To nHide)
                Set hid(nHide) = b
                nHide = nHide + 1
            End If
        Next i
    End If
    If nHide = 0 Or nKeep = 0 Then
        ' nothing hidden in the assembly (or no name matched): whole model as before
        For f = 0 To gFormatCount - 1
            If gFormats(f).Selected And gFormats(f).Kind = "MODEL" Then
                SaveOne m, gFormats(f), cfgName, "", ""
            End If
        Next f
        Exit Sub
    End If

    On Error GoTo EH
    For i = 0 To nHide - 1
        hid(i).HideBody True
    Next i
    m.GraphicsRedraw2
    For f = 0 To gFormatCount - 1
        If gFormats(f).Selected And gFormats(f).Kind = "MODEL" And gFormats(f).Flag <> "ISOLATE" Then
            SaveOne m, gFormats(f), cfgName, "", ""
        End If
    Next f
    vKeep = keep
    SaveIsolatedBodies m, vKeep, cfgName, ""
Cleanup:
    On Error Resume Next
    For i = 0 To nHide - 1
        hid(i).HideBody False
    Next i
    m.GraphicsRedraw2
    Exit Sub
EH:
    AddLog DocKey(m), cfgName, "", "", "", "", "ERROR: " & Err.Description, 0, 0
    mFail = mFail + 1
    Resume Cleanup
End Sub

' Copies one body into a temporary (hidden) part, saves it in every ticked ISOLATE
' format, then closes the temporary part without saving.
Private Sub SaveIsolated(src As SldWorks.ModelDoc2, ByVal body As SldWorks.Body2, _
                         ByVal cfgName As String, ByVal bName As String)
    Dim one() As Object, v As Variant
    ReDim one(0 To 0)
    Set one(0) = body
    v = one
    SaveIsolatedBodies src, v, cfgName, bName
End Sub

' Same, for one or more bodies (Variant array of Body2): all go into one temporary part.
Private Sub SaveIsolatedBodies(src As SldWorks.ModelDoc2, ByVal bodies As Variant, _
                               ByVal cfgName As String, ByVal bName As String)
    Dim f As Long, nTodo As Long, outPath As String, paths() As String, i As Long
    Dim tmpl As String, tmp As SldWorks.ModelDoc2, tmpPart As SldWorks.PartDoc, body As SldWorks.Body2
    Dim feat As SldWorks.Feature, e As Long, w As Long, ok As Boolean, visOff As Boolean

    If Not IsArray(bodies) Then Exit Sub
    ReDim paths(0 To gFormatCount)
    For f = 0 To gFormatCount - 1
        If gFormats(f).Selected And gFormats(f).Kind = "MODEL" And gFormats(f).Flag = "ISOLATE" Then
            outPath = PlanPath(src, gFormats(f), cfgName, bName, "")
            If outPath <> "" Then
                If gDryRun Then
                    LogResult src, cfgName, bName, "", gFormats(f), outPath, "DRY RUN", 0, 0
                Else
                    paths(f) = outPath
                    nTodo = nTodo + 1
                End If
            End If
        End If
    Next f
    If nTodo = 0 Then Exit Sub

    On Error GoTo EH
    tmpl = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
    If tmpl <> "" Then
        visOff = True
        swApp.DocumentVisible False, swDocumentTypes_e.swDocPART
        Set tmp = swApp.NewDocument(tmpl, 0, 0, 0)
        swApp.DocumentVisible True, swDocumentTypes_e.swDocPART
        visOff = False
    End If
    If tmp Is Nothing Then
        FailPaths src, cfgName, bName, paths, "FAILED: could not create a temporary part (set Tools > Options > Default Templates > Parts)"
        GoTo Cleanup
    End If
    Set tmpPart = tmp
    For i = LBound(bodies) To UBound(bodies)
        Set body = bodies(i)
        Set feat = Nothing
        Set feat = tmpPart.CreateFeatureFromBody3(body.Copy, False, swCreateFeatureBodyOpts_e.swCreateFeatureBodySimplify)
        If feat Is Nothing Then Set feat = tmpPart.CreateFeatureFromBody3(body.Copy, False, swCreateFeatureBodyOpts_e.swCreateFeatureBodyCheck)
        If feat Is Nothing Then Set feat = tmpPart.CreateFeatureFromBody3(body.Copy, False, 0)
        If feat Is Nothing Then
            FailPaths src, cfgName, bName, paths, "FAILED: CreateFeatureFromBody3 could not copy the body"
            GoTo Cleanup
        End If
    Next i
    For f = 0 To gFormatCount - 1
        If paths(f) <> "" Then
            e = 0: w = 0
            ok = tmp.Extension.SaveAs3(paths(f), swSaveAsVersion_e.swSaveAsCurrentVersion, _
                    swSaveAsOptions_e.swSaveAsOptions_Silent Or swSaveAsOptions_e.swSaveAsOptions_Copy, _
                    Nothing, Nothing, e, w)
            If ok Then ok = FS().FileExists(paths(f))
            LogResult src, cfgName, bName, "", gFormats(f), paths(f), IIf(ok, "OK", "FAILED"), e, w
            paths(f) = ""
        End If
    Next f
Cleanup:
    On Error Resume Next
    If visOff Then swApp.DocumentVisible True, swDocumentTypes_e.swDocPART
    If Not tmp Is Nothing Then swApp.CloseDoc tmp.GetTitle      ' never saved
    swApp.ActivateDoc3 DocKey(src), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, e
    Exit Sub
EH:
    FailPaths src, cfgName, bName, paths, "ERROR: " & Err.Description
    Resume Cleanup
End Sub

Private Sub FailPaths(src As SldWorks.ModelDoc2, ByVal cfgName As String, ByVal bName As String, _
                      paths() As String, ByVal msg As String)
    Dim f As Long
    For f = 0 To gFormatCount - 1
        If paths(f) <> "" Then
            LogResult src, cfgName, bName, "", gFormats(f), paths(f), msg, 0, 0
            paths(f) = ""
        End If
    Next f
End Sub

'---------------------------- drawings -----------------------------------------
Private Sub ExportDrawing(m As SldWorks.ModelDoc2, ByVal sheets As Collection)
    Dim drw As SldWorks.DrawingDoc, origSheet As String, f As Long, i As Long, n As Long
    Dim names() As String, one(0 To 0) As String, vAll As Variant, vOne As Variant

    If sheets Is Nothing Then Exit Sub
    n = sheets.Count
    If n = 0 Then Exit Sub
    Set drw = m
    origSheet = drw.GetCurrentSheet.GetName
    mModelMode = MODE_SHEETS
    ReDim names(0 To n - 1)
    For i = 1 To n
        names(i - 1) = CStr(sheets(i))
    Next i
    vAll = names

    On Error GoTo EH
    For f = 0 To gFormatCount - 1
        If gCancel Then Exit For
        If gFormats(f).Selected And gFormats(f).Kind = "DRAWING" Then
            If gFormats(f).Ext = ".pdf" Then
                If gPdfCombined Then
                    ' one PDF with just the ticked sheets
                    SaveOne m, gFormats(f), "", "", "", NewPdfData(vAll)
                Else
                    ' one PDF per sheet, each with its own name
                    For i = 0 To n - 1
                        If gCancel Then Exit For
                        one(0) = names(i)
                        vOne = one
                        SaveOne m, gFormats(f), "", "", names(i), NewPdfData(vOne)
                    Next i
                End If
            ElseIf gFormats(f).Flag = "SINGLE" Then
                SaveOne m, gFormats(f), "", "", ""
            Else
                ' one file per sheet (DXF/DWG use "active sheet only" for this run)
                For i = 0 To n - 1
                    If gCancel Then Exit For
                    If drw.ActivateSheet(names(i)) Then
                        m.GraphicsRedraw2
                        SaveOne m, gFormats(f), "", "", names(i)
                    Else
                        AddLog DocKey(m), "", "", names(i), gFormats(f).Label, "", "FAILED: could not activate sheet", 0, 0
                        mFail = mFail + 1
                    End If
                Next i
            End If
        End If
    Next f
Cleanup:
    On Error Resume Next
    If drw.GetCurrentSheet.GetName <> origSheet Then drw.ActivateSheet origSheet
    Exit Sub
EH:
    AddLog DocKey(m), "", "", "", "", "", "ERROR: " & Err.Description, 0, 0
    mFail = mFail + 1
    Resume Cleanup
End Sub

Private Function NewPdfData(ByVal sheetNames As Variant) As Object
    Dim pd As Object
    Set pd = swApp.GetExportFileData(swExportDataFileType_e.swExportPdfData)
    If Not pd Is Nothing Then
        pd.SetSheets swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, sheetNames
    End If
    Set NewPdfData = pd
End Function

Private Function Make3DPdfData() As Object
    Dim pd As Object                 ' late bound on purpose
    On Error Resume Next
    Set pd = swApp.GetExportFileData(swExportDataFileType_e.swExportPdfData)
    If Not pd Is Nothing Then pd.ExportAs3D = True
    Set Make3DPdfData = pd
End Function

'---------------------------- one SaveAs ---------------------------------------
Private Sub SaveOne(src As SldWorks.ModelDoc2, fmt As ExportFormat, ByVal cfgName As String, _
                    ByVal bName As String, ByVal sheet As String, Optional ByVal expData As Object = Nothing)
    Dim outPath As String, e As Long, w As Long, ok As Boolean, errMsg As String
    Dim ext As String, how As String, writerErr As String
    outPath = PlanPath(src, fmt, cfgName, bName, sheet)
    If outPath = "" Then Exit Sub
    If gDryRun Then
        LogResult src, cfgName, bName, sheet, fmt, outPath, "DRY RUN", 0, 0
        Exit Sub
    End If
    ext = LCase$(fmt.Ext)
    ' A part/assembly saved as .pdf is a 2D picture unless ExportAs3D is set
    If expData Is Nothing And ext = ".pdf" And src.GetType <> swDocumentTypes_e.swDocDRAWING Then
        Set expData = Make3DPdfData()
    End If

    ' 1) SOLIDWORKS Save As
    On Error Resume Next
    ok = src.Extension.SaveAs3(outPath, swSaveAsVersion_e.swSaveAsCurrentVersion, _
            swSaveAsOptions_e.swSaveAsOptions_Silent Or swSaveAsOptions_e.swSaveAsOptions_Copy, _
            expData, Nothing, e, w)
    If Err.Number <> 0 Then
        ok = False
        errMsg = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' 2) GLB / OBJ: write the file from the model's display tessellation
    If Not ok And (ext = ".glb" Or ext = ".obj") Then
        ok = WriteMeshFile(src, outPath, ext, writerErr)
        If ok Then
            how = " (mesh writer)"
            e = 0: w = 0: errMsg = ""
        End If
    End If

    ' 3) fallbacks: Save As without the Copy option, then the older SaveAs4
    If Not ok Then
        On Error Resume Next
        e = 0: w = 0
        ok = src.Extension.SaveAs3(outPath, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                swSaveAsOptions_e.swSaveAsOptions_Silent, expData, Nothing, e, w)
        If Err.Number <> 0 Then
            ok = False
            errMsg = Err.Description
            Err.Clear
        End If
        If Not ok Then
            e = 0: w = 0
            ok = src.SaveAs4(outPath, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                    swSaveAsOptions_e.swSaveAsOptions_Silent, e, w)
            If Err.Number <> 0 Then
                ok = False
                errMsg = Err.Description
                Err.Clear
            End If
        End If
        On Error GoTo 0
    End If

    If ok And Not FS().FileExists(outPath) Then
        ok = False
        errMsg = "SaveAs reported success but no file was written"
    End If
    If Not ok And writerErr <> "" Then errMsg = errMsg & IIf(errMsg <> "", " / ", "") & "mesh writer: " & writerErr
    LogResult src, cfgName, bName, sheet, fmt, outPath, IIf(ok, "OK" & how, "FAILED" & IIf(errMsg <> "", ": " & errMsg, "")), e, w
End Sub

Private Sub LogResult(src As SldWorks.ModelDoc2, ByVal cfgName As String, ByVal bName As String, _
                      ByVal sheet As String, fmt As ExportFormat, ByVal outPath As String, _
                      ByVal status As String, ByVal e As Long, ByVal w As Long)
    AddLog DocKey(src), cfgName, bName, sheet, fmt.Label, outPath, status, e, w
    If Left$(status, 2) = "OK" Or status = "DRY RUN" Then
        mOk = mOk + 1
    Else
        mFail = mFail + 1
    End If
    mStep = mStep + 1
    SetStatus "SWExporter: " & mStep & " done - " & Mid$(outPath, InStrRev(outPath, "\") + 1)
    DoEvents
End Sub

Private Sub AddLog(ByVal src As String, ByVal cfgName As String, ByVal bName As String, ByVal sheet As String, _
                   ByVal fmtLabel As String, ByVal outPath As String, ByVal status As String, _
                   ByVal e As Long, ByVal w As Long)
    mLog = mLog & CsvRow(Now, src, cfgName, bName, sheet, fmtLabel, outPath, status, e, w)
End Sub

'=========================== GLB / OBJ MESH WRITER ============================
' Builds the file from Face2.GetTessTriangles / GetTessNorms of every visible body
' (a part's visible bodies, or the visible part components of an assembly).
Private Function WriteMeshFile(src As SldWorks.ModelDoc2, ByVal path As String, ByVal ext As String, why As String) As Boolean
    On Error GoTo EH
    If Not CollectMesh(src) Then
        why = "no tessellation found (nothing visible?)"
        GoTo Done
    End If
    If ext = ".glb" Then
        WriteMeshFile = WriteGlb(path)
    Else
        WriteMeshFile = WriteObj(path)
    End If
    If WriteMeshFile Then WriteMeshFile = FS().FileExists(path)
    If Not WriteMeshFile And why = "" Then why = "could not write the file"
Done:
    Erase mPos
    Erase mNrm
    mCnt = 0
    Exit Function
EH:
    why = Err.Description
    WriteMeshFile = False
    Resume Done
End Function

Private Function CollectMesh(src As SldWorks.ModelDoc2) As Boolean
    Dim v As Variant, i As Long
    MeshReset
    Select Case src.GetType
        Case swDocumentTypes_e.swDocPART
            v = VisibleBodies(src, True)
            If IsArray(v) Then
                For i = LBound(v) To UBound(v)
                    MeshAddBody v(i)
                Next i
            End If
        Case swDocumentTypes_e.swDocASSEMBLY
            MeshAddAssembly src
    End Select
    CollectMesh = (mCnt >= 9)
End Function

Private Sub MeshReset()
    mCnt = 0
    ReDim mPos(0 To 89999)
    ReDim mNrm(0 To 89999)
End Sub

Private Sub MeshEnsure(ByVal need As Long)
    Dim n As Long
    If need - 1 <= UBound(mPos) Then Exit Sub
    n = UBound(mPos) + 1
    Do While n < need
        n = n * 2
    Loop
    ReDim Preserve mPos(0 To n - 1)
    ReDim Preserve mNrm(0 To n - 1)
End Sub

Private Sub MeshAddBody(ByVal b As Object)
    Dim vF As Variant, i As Long, f As Object, t As Variant, nr As Variant
    Dim k As Long, n As Long, lb As Long, nlb As Long, useN As Boolean
    vF = b.GetFaces
    If Not IsArray(vF) Then Exit Sub
    For i = LBound(vF) To UBound(vF)
        Set f = vF(i)
        t = Empty
        t = f.GetTessTriangles(True)                 ' True = meters, no unit conversion
        If IsArray(t) Then
            n = UBound(t) - LBound(t) + 1
            n = n - (n Mod 9)
            If n > 0 Then
                nr = Empty
                nr = f.GetTessNorms
                useN = False
                If IsArray(nr) Then useN = ((UBound(nr) - LBound(nr) + 1) >= n)
                MeshEnsure mCnt + n
                lb = LBound(t)
                If useN Then nlb = LBound(nr)
                For k = 0 To n - 1
                    mPos(mCnt + k) = t(lb + k)
                    If useN Then mNrm(mCnt + k) = nr(nlb + k)
                Next k
                If Not useN Then FlatNormals mCnt, n
                mCnt = mCnt + n
            End If
        End If
    Next i
End Sub

Private Sub FlatNormals(ByVal s0 As Long, ByVal n As Long)
    Dim k As Long, j As Long, ux As Double, uy As Double, uz As Double
    Dim vx As Double, vy As Double, vz As Double, nx As Double, ny As Double, nz As Double, l As Double
    For k = s0 To s0 + n - 9 Step 9
        ux = mPos(k + 3) - mPos(k): uy = mPos(k + 4) - mPos(k + 1): uz = mPos(k + 5) - mPos(k + 2)
        vx = mPos(k + 6) - mPos(k): vy = mPos(k + 7) - mPos(k + 1): vz = mPos(k + 8) - mPos(k + 2)
        nx = uy * vz - uz * vy
        ny = uz * vx - ux * vz
        nz = ux * vy - uy * vx
        l = Sqr(nx * nx + ny * ny + nz * nz)
        If l > 0 Then
            nx = nx / l: ny = ny / l: nz = nz / l
        Else
            nx = 0: ny = 0: nz = 1
        End If
        For j = 0 To 6 Step 3
            mNrm(k + j) = nx: mNrm(k + j + 1) = ny: mNrm(k + j + 2) = nz
        Next j
    Next k
End Sub

Private Sub MeshAddAssembly(asm As SldWorks.ModelDoc2)
    Dim a As Object, v As Variant, i As Long, j As Long, comp As Object
    Dim vb As Variant, info As Variant, s0 As Long, vis As Boolean
    Set a = asm
    v = a.GetComponents(False)
    If Not IsArray(v) Then Exit Sub
    For i = LBound(v) To UBound(v)
        Set comp = v(i)
        If UsablePartComp(comp, True) Then
            s0 = mCnt
            vb = Empty
            On Error Resume Next
            vb = comp.GetBodies3(swBodyType_e.swSolidBody, info)
            On Error GoTo 0
            If IsArray(vb) Then
                For j = LBound(vb) To UBound(vb)
                    vis = True
                    On Error Resume Next
                    vis = vb(j).Visible
                    On Error GoTo 0
                    If vis Then MeshAddBody vb(j)
                Next j
            End If
            If mCnt > s0 Then PlaceComponent comp, s0
        End If
    Next i
End Sub

' Moves the component's triangles into assembly space with Component2.Transform2.
' If the tessellation already sits where Component2.GetBox says the component is,
' it is left alone, so the transform is never applied twice.
Private Sub PlaceComponent(ByVal comp As Object, ByVal s0 As Long)
    Dim xf As Object, a As Variant, box As Variant, k As Long, j As Long
    Dim mn(0 To 2) As Double, mx(0 To 2) As Double, c0(0 To 2) As Double, ct(0 To 2) As Double, bc(0 To 2) As Double
    Dim x As Double, y As Double, z As Double, sc As Double, dRaw As Double, dXf As Double
    On Error Resume Next
    Set xf = comp.Transform2
    If xf Is Nothing Then Exit Sub
    a = xf.ArrayData
    If Not IsArray(a) Then Exit Sub
    If UBound(a) < 12 Then Exit Sub
    sc = a(12)
    If sc = 0 Then sc = 1

    For j = 0 To 2
        mn(j) = 1E+30: mx(j) = -1E+30
    Next j
    For k = s0 To mCnt - 1
        j = (k - s0) Mod 3
        If mPos(k) < mn(j) Then mn(j) = mPos(k)
        If mPos(k) > mx(j) Then mx(j) = mPos(k)
    Next k
    For j = 0 To 2
        c0(j) = (mn(j) + mx(j)) / 2
    Next j
    ct(0) = sc * (c0(0) * a(0) + c0(1) * a(3) + c0(2) * a(6)) + a(9)
    ct(1) = sc * (c0(0) * a(1) + c0(1) * a(4) + c0(2) * a(7)) + a(10)
    ct(2) = sc * (c0(0) * a(2) + c0(1) * a(5) + c0(2) * a(8)) + a(11)

    box = Empty
    box = comp.GetBox(False, False)
    If IsArray(box) Then
        If UBound(box) - LBound(box) >= 5 Then
            For j = 0 To 2
                bc(j) = (box(LBound(box) + j) + box(LBound(box) + j + 3)) / 2
            Next j
            dRaw = (c0(0) - bc(0)) ^ 2 + (c0(1) - bc(1)) ^ 2 + (c0(2) - bc(2)) ^ 2
            dXf = (ct(0) - bc(0)) ^ 2 + (ct(1) - bc(1)) ^ 2 + (ct(2) - bc(2)) ^ 2
            If dRaw < dXf Then Exit Sub                 ' already in assembly space
        End If
    End If

    For k = s0 To mCnt - 1 Step 3
        x = mPos(k): y = mPos(k + 1): z = mPos(k + 2)
        mPos(k) = sc * (x * a(0) + y * a(3) + z * a(6)) + a(9)
        mPos(k + 1) = sc * (x * a(1) + y * a(4) + z * a(7)) + a(10)
        mPos(k + 2) = sc * (x * a(2) + y * a(5) + z * a(8)) + a(11)
        x = mNrm(k): y = mNrm(k + 1): z = mNrm(k + 2)
        mNrm(k) = x * a(0) + y * a(3) + z * a(6)
        mNrm(k + 1) = x * a(1) + y * a(4) + z * a(7)
        mNrm(k + 2) = x * a(2) + y * a(5) + z * a(8)
    Next k
End Sub

' Binary glTF 2.0: one mesh, POSITION + NORMAL (float32), meters, one grey material.
Private Function WriteGlb(ByVal path As String) As Boolean
    Dim fn As Integer, nV As Long, half As Long, binLen As Long, json As String, pad As Long
    Dim mn(0 To 2) As Single, mx(0 To 2) As Single, k As Long, j As Long
    Dim hMagic As Long, hVer As Long, hLen As Long, cLen As Long, cType As Long

    If mCnt < 9 Then Exit Function
    ReDim Preserve mPos(0 To mCnt - 1)
    ReDim Preserve mNrm(0 To mCnt - 1)
    nV = mCnt \ 3
    For j = 0 To 2
        mn(j) = mPos(j): mx(j) = mPos(j)
    Next j
    For k = 0 To mCnt - 1
        j = k Mod 3
        If mPos(k) < mn(j) Then mn(j) = mPos(k)
        If mPos(k) > mx(j) Then mx(j) = mPos(k)
    Next k
    half = mCnt * 4                    ' bytes per attribute block
    binLen = half * 2

    json = "{""asset"":{""version"":""2.0"",""generator"":""SWExporter""}," & _
           """scene"":0,""scenes"":[{""nodes"":[0]}],""nodes"":[{""mesh"":0}]," & _
           """meshes"":[{""primitives"":[{""attributes"":{""POSITION"":0,""NORMAL"":1},""material"":0}]}]," & _
           """materials"":[{""pbrMetallicRoughness"":{""baseColorFactor"":[0.8,0.8,0.8,1],""metallicFactor"":0.1,""roughnessFactor"":0.6},""doubleSided"":true}]," & _
           """buffers"":[{""byteLength"":" & binLen & "}]," & _
           """bufferViews"":[{""buffer"":0,""byteOffset"":0,""byteLength"":" & half & ",""target"":34962}," & _
           "{""buffer"":0,""byteOffset"":" & half & ",""byteLength"":" & half & ",""target"":34962}]," & _
           """accessors"":[{""bufferView"":0,""componentType"":5126,""count"":" & nV & ",""type"":""VEC3""," & _
           """min"":[" & Num(mn(0)) & "," & Num(mn(1)) & "," & Num(mn(2)) & "]," & _
           """max"":[" & Num(mx(0)) & "," & Num(mx(1)) & "," & Num(mx(2)) & "]}," & _
           "{""bufferView"":1,""componentType"":5126,""count"":" & nV & ",""type"":""VEC3""}]}"
    pad = (4 - (Len(json) Mod 4)) Mod 4
    json = json & Space$(pad)

    On Error GoTo EH
    If FS().FileExists(path) Then FS().DeleteFile path, True    ' Open For Binary does not truncate
    fn = FreeFile
    Open path For Binary Access Write As #fn
    hMagic = &H46546C67                 ' "glTF"
    hVer = 2
    hLen = 12 + 8 + Len(json) + 8 + binLen
    Put #fn, , hMagic
    Put #fn, , hVer
    Put #fn, , hLen
    cLen = Len(json)
    cType = &H4E4F534A                  ' "JSON"
    Put #fn, , cLen
    Put #fn, , cType
    Put #fn, , json
    cLen = binLen
    cType = &H4E4942                    ' "BIN"
    Put #fn, , cLen
    Put #fn, , cType
    Put #fn, , mPos
    Put #fn, , mNrm
    Close #fn
    fn = 0
    WriteGlb = True
    Exit Function
EH:
    If fn <> 0 Then Close #fn
    WriteGlb = False
End Function

' Wavefront OBJ: v / vn / f, in the STL units chosen in the form.
Private Function WriteObj(ByVal path As String) As Boolean
    Dim ts As Object, sc As Double, k As Long, buf() As String, n As Long, a As Long, nTri As Long
    sc = UnitScale(gStlUnits)
    On Error GoTo EH
    Set ts = FS().CreateTextFile(path, True, False)
    ts.WriteLine "# SWExporter OBJ, units: " & UnitName(gStlUnits)
    ReDim buf(0 To 4999)
    For k = 0 To mCnt - 1 Step 3
        buf(n) = "v " & Num(mPos(k) * sc) & " " & Num(mPos(k + 1) * sc) & " " & Num(mPos(k + 2) * sc)
        n = n + 1
        If n > UBound(buf) Then FlushLines ts, buf, n
    Next k
    For k = 0 To mCnt - 1 Step 3
        buf(n) = "vn " & Num(mNrm(k)) & " " & Num(mNrm(k + 1)) & " " & Num(mNrm(k + 2))
        n = n + 1
        If n > UBound(buf) Then FlushLines ts, buf, n
    Next k
    nTri = mCnt \ 9
    For k = 0 To nTri - 1
        a = k * 3 + 1
        buf(n) = "f " & a & "//" & a & " " & (a + 1) & "//" & (a + 1) & " " & (a + 2) & "//" & (a + 2)
        n = n + 1
        If n > UBound(buf) Then FlushLines ts, buf, n
    Next k
    FlushLines ts, buf, n
    ts.Close
    WriteObj = True
    Exit Function
EH:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    WriteObj = False
End Function

Private Sub FlushLines(ByVal ts As Object, buf() As String, n As Long)
    If n = 0 Then Exit Sub
    If n - 1 < UBound(buf) Then ReDim Preserve buf(0 To n - 1)
    ts.WriteLine Join(buf, vbCrLf)
    ReDim buf(0 To 4999)
    n = 0
End Sub

' Number with "." as decimal separator whatever the Windows locale (JSON/OBJ safe).
Private Function Num(ByVal x As Double) As String
    Dim s As String
    s = Trim$(Str$(CSng(x)))
    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & Mid$(s, 2)
    End If
    Num = s
End Function

Private Function UnitScale(ByVal u As Long) As Double
    Select Case u
        Case 0: UnitScale = 1000#
        Case 1: UnitScale = 100#
        Case 3: UnitScale = 39.3700787401575
        Case 4: UnitScale = 3.28083989501312
        Case Else: UnitScale = 1#
    End Select
End Function

Private Function UnitName(ByVal u As Long) As String
    UnitName = Choose(u + 1, "mm", "cm", "m", "in", "ft")
    If UnitName = "" Then UnitName = "m"
End Function

'---------------------------- naming / paths -----------------------------------
' Returns "" when the file should not be written (the reason is logged).
Private Function PlanPath(src As SldWorks.ModelDoc2, fmt As ExportFormat, ByVal cfgName As String, _
                          ByVal bName As String, ByVal sheet As String) As String
    Dim outDir As String, base As String, ext As String, tmpl As String, p As String, n As Long
    outDir = OutputDir(src)
    If outDir = "" Then Exit Function

    ext = fmt.Ext
    If ext = ".e*" Then
        Select Case src.GetType
            Case swDocumentTypes_e.swDocPART: ext = ".eprt"
            Case swDocumentTypes_e.swDocASSEMBLY: ext = ".easm"
            Case Else: ext = ".edrw"
        End Select
    End If

    tmpl = gTemplate
    If tmpl = "" Then
        If sheet <> "" Then
            tmpl = "{file}_{sheet}"
        ElseIf bName <> "" Then
            tmpl = "{file}_{config}_{body}"
        Else
            tmpl = "{file}_{config}"
        End If
        ' leave the config out for drawings, single-config models and Bodies mode (also in batch)
        If src.GetType = swDocumentTypes_e.swDocDRAWING Or src.GetConfigurationCount <= 1 _
           Or mModelMode = MODE_BODIES Then
            tmpl = Replace(tmpl, "_{config}", "")
        End If
    End If
    base = tmpl
    base = Replace(base, "{file}", DocName(src))
    base = Replace(base, "{component}", mCompName)
    base = Replace(base, "{config}", cfgName)
    base = Replace(base, "{body}", bName)
    base = Replace(base, "{sheet}", sheet)
    base = Replace(base, "{date}", Format$(Date, "yyyy-mm-dd"))
    If InStr(1, base, "{rev}", vbTextCompare) > 0 Then base = Replace(base, "{rev}", GetRevision(src, cfgName), , , vbTextCompare)
    base = CleanName(base)

    p = outDir & "\" & base & ext
    If FileOrPlanned(p) Then
        Select Case gCollision
            Case 1      ' skip
                AddLog DocKey(src), cfgName, bName, sheet, fmt.Label, p, "SKIPPED (exists)", 0, 0
                mSkip = mSkip + 1
                Exit Function
            Case 2      ' overwrite - but never our own output from this run
                If InCollectionKey(mPlanned, LCase$(p)) Then
                    n = 2
                    Do While FileOrPlanned(outDir & "\" & base & "_" & n & ext)
                        n = n + 1
                    Loop
                    p = outDir & "\" & base & "_" & n & ext
                End If
            Case Else   ' auto-number
                n = 2
                Do While FileOrPlanned(outDir & "\" & base & "_" & n & ext)
                    n = n + 1
                Loop
                p = outDir & "\" & base & "_" & n & ext
        End Select
    End If
    mPlanned.Add p, LCase$(p)
    PlanPath = p
End Function

Private Function FileOrPlanned(ByVal p As String) As Boolean
    If InCollectionKey(mPlanned, LCase$(p)) Then
        FileOrPlanned = True
    ElseIf gDryRun And gCollision = 2 Then
        FileOrPlanned = False
    Else
        On Error Resume Next
        FileOrPlanned = FS().FileExists(p)
    End If
End Function

' Output folder. While exporting the parts of an assembly, it follows the assembly
' (next to the assembly, or a subfolder named after it).
Private Function OutputDir(src As SldWorks.ModelDoc2) As String
    Dim anchor As SldWorks.ModelDoc2, p As String, base As String, subName As String, d As String
    If mAnchorDoc Is Nothing Then Set anchor = src Else Set anchor = mAnchorDoc
    p = anchor.GetPathName
    If p <> "" Then base = Left$(p, InStrRev(p, "\") - 1)
    If base <> "" And mFirstSrcDir = "" Then mFirstSrcDir = base
    If base = "" And gOutMode <> 2 Then
        AddLog anchor.GetTitle, "", "", "", "", "", "SKIPPED: document has never been saved - use a custom output folder", 0, 0
        mSkip = mSkip + 1
        Exit Function
    End If
    Select Case gOutMode
        Case 1
            subName = gOutText
            If subName = "" Then subName = "{file}_export"
            subName = Replace(subName, "{file}", DocName(anchor))
            subName = Replace(subName, "{date}", Format$(Date, "yyyy-mm-dd"))
            d = base & "\" & CleanName(subName)
        Case 2
            d = Trim$(gOutText)
            Do While Len(d) > 3 And Right$(d, 1) = "\"
                d = Left$(d, Len(d) - 1)
            Loop
        Case Else
            d = base
    End Select
    If Not gDryRun Then
        If Not EnsureFolder(d) Then
            AddLog DocKey(src), "", "", "", "", d, "FAILED: could not create output folder", 0, 0
            mFail = mFail + 1
            Exit Function
        End If
    End If
    If mFirstOutDir = "" Then mFirstOutDir = d
    OutputDir = d
End Function

Public Function DocName(m As SldWorks.ModelDoc2) As String
    Dim p As String
    p = m.GetPathName
    If p = "" Then p = m.GetTitle
    p = Mid$(p, InStrRev(p, "\") + 1)
    If InStrRev(p, ".") > 0 Then
        Select Case LCase$(Mid$(p, InStrRev(p, ".") + 1))
            Case "sldprt", "sldasm", "slddrw": p = Left$(p, InStrRev(p, ".") - 1)
        End Select
    End If
    DocName = p
End Function

' Full path when saved, otherwise the title. Use this for ActivateDoc3 / CloseDoc:
' a drawing's GetTitle includes " - Sheet1" and is not reliable there.
Public Function DocKey(m As SldWorks.ModelDoc2) As String
    Dim k As String
    On Error Resume Next
    k = m.GetPathName
    If k = "" Then k = m.GetTitle
    DocKey = k
End Function

Private Function GetRevision(m As SldWorks.ModelDoc2, ByVal cfgName As String) As String
    Dim cpm As SldWorks.CustomPropertyManager, v As String, r As String, nm As Variant
    On Error Resume Next
    For Each nm In Array("Revision", "Rev")
        If cfgName <> "" Then
            Set cpm = m.Extension.CustomPropertyManager(cfgName)
            v = "": r = ""
            If Not cpm Is Nothing Then cpm.Get4 CStr(nm), False, v, r
            If r <> "" Then GetRevision = r: Exit Function
        End If
        Set cpm = m.Extension.CustomPropertyManager("")
        v = "": r = ""
        If Not cpm Is Nothing Then cpm.Get4 CStr(nm), False, v, r
        If r <> "" Then GetRevision = r: Exit Function
    Next nm
End Function

' Characters Windows doesn't allow in file names become "_"; leftover separators are tidied.
Public Function CleanName(ByVal s As String) As String
    Dim bad As Variant, k As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For k = LBound(bad) To UBound(bad)
        s = Replace(s, bad(k), "_")
    Next k
    For k = 1 To 31
        s = Replace(s, Chr$(k), "_")
    Next k
    Do While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Loop
    Do While InStr(s, "_-_") > 0
        s = Replace(s, "_-_", "_")
    Loop
    s = Trim$(s)
    Do While Len(s) > 0
        If InStr("._- ", Right$(s, 1)) = 0 Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    Do While Len(s) > 0
        If InStr("_- ", Left$(s, 1)) = 0 Then Exit Do
        s = Mid$(s, 2)
    Loop
    If s = "" Then s = "export"
    If Len(s) > 180 Then s = Left$(s, 180)
    CleanName = s
End Function

'---------------------------- selection save/restore ---------------------------
Private Function SaveSelection(m As SldWorks.ModelDoc2) As Collection
    Dim c As Collection, sm As SldWorks.SelectionMgr, i As Long, n As Long
    Set c = New Collection
    On Error Resume Next
    Set sm = m.SelectionManager
    n = sm.GetSelectedObjectCount2(-1)
    For i = 1 To n
        c.Add sm.GetSelectedObject6(i, -1)
    Next i
    Set SaveSelection = c
End Function

Private Sub RestoreSelection(m As SldWorks.ModelDoc2, ByVal c As Collection)
    Dim i As Long, o As Object, ent As SldWorks.Entity, feat As SldWorks.Feature
    On Error Resume Next
    m.ClearSelection2 True
    If c Is Nothing Then Exit Sub
    For i = 1 To c.Count
        Set o = Nothing
        Set o = c(i)
        If Not o Is Nothing Then
            Set ent = Nothing
            Set ent = o
            If Not ent Is Nothing Then
                ent.Select4 True, Nothing
            Else
                Set feat = Nothing
                Set feat = o
                If Not feat Is Nothing Then feat.Select2 True, -1
            End If
        End If
    Next i
End Sub

'---------------------------- small utilities ----------------------------------
Public Sub SetStatus(ByVal s As String)
    Dim fr As SldWorks.Frame
    On Error Resume Next
    Set fr = swApp.Frame
    If Not fr Is Nothing Then fr.SetStatusBarText s
    If gFormLoaded Then frmSWExporter.ShowStatus s     ' never auto-load the form from here
End Sub

Public Function InCollection(ByVal c As Collection, ByVal s As String) As Boolean
    ' Nothing = "everything"
    Dim i As Long
    If c Is Nothing Then
        InCollection = True
        Exit Function
    End If
    For i = 1 To c.Count
        If CStr(c(i)) = s Then
            InCollection = True
            Exit Function
        End If
    Next i
End Function

Private Function InCollectionKey(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    If c Is Nothing Then Exit Function
    On Error Resume Next
    v = c(key)
    InCollectionKey = (Err.Number = 0)
    Err.Clear
End Function

Private Function CsvRow(ByVal t As Date, ByVal a As String, ByVal b As String, ByVal c As String, _
                        ByVal d As String, ByVal e As String, ByVal f As String, ByVal g As String, _
                        ByVal errs As Long, ByVal warns As Long) As String
    CsvRow = Format$(t, "yyyy-mm-dd hh:nn:ss") & "," & Q(a) & "," & Q(b) & "," & Q(c) & "," & Q(d) & "," & _
             Q(e) & "," & Q(f) & "," & Q(g) & "," & errs & "," & warns & vbCrLf
End Function

Private Function Q(ByVal s As String) As String
    Q = """" & Replace(s, """", """""") & """"
End Function

Public Function EnsureFolder(ByVal p As String) As Boolean
    Dim parent As String
    On Error Resume Next
    If p = "" Then Exit Function
    If Not FS().FolderExists(p) Then
        parent = FS().GetParentFolderName(p)
        If parent <> "" Then
            If Not FS().FolderExists(parent) Then EnsureFolder parent
        End If
        FS().CreateFolder p
    End If
    EnsureFolder = FS().FolderExists(p)
End Function

' Writes ANSI; falls back to Unicode if a name in the log has characters ANSI can't hold.
Private Function WriteText(ByVal p As String, ByVal s As String) As Boolean
    Dim ts As Object
    On Error Resume Next
    Set ts = FS().CreateTextFile(p, True, False)
    ts.Write s
    If Err.Number <> 0 Then
        Err.Clear
        If Not ts Is Nothing Then ts.Close
        Set ts = Nothing
        Set ts = FS().CreateTextFile(p, True, True)
        ts.Write s
    End If
    ts.Close
    WriteText = (Err.Number = 0)
    Err.Clear
    If WriteText Then WriteText = FS().FileExists(p)
End Function

Public Function BrowseForFolder(ByVal prompt As String) As String
    Dim sh As Object, fld As Object
    On Error Resume Next
    Set sh = CreateObject("Shell.Application")
    Set fld = sh.BrowseForFolder(0, prompt, &H51)   ' file-system folders only, edit box, new dialog style
    If fld Is Nothing Then Exit Function
    BrowseForFolder = fld.Self.Path
End Function
