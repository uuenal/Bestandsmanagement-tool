Attribute VB_Name = "SHIELD_Tool"
Option Explicit
' =====================================================================
'  SHIELD prototype material check
'
'  Sheet "Start": enter the assembly number + units to build, then click
'     1 Update BOM         -> sheet "BOM"        (CAD structure from Windchill, read only)
'     2 Update SAP         -> sheet "SAP"        (LX02 via SAP GUI scripting)
'     3 Update order list  -> sheet "Order list" (SharePoint file, read only)
'     4 Update CSP Hall    -> sheet "CSP Hall"   (copy of SHIELD_CSP_Hall.xlsx next to this file)
'     5 Warehouse request  -> sheet "Auslagerung" + Outlook draft to Samet, logged in the CSP Hall file
'  After every button the sheet "Status" is rebuilt: one row per part
'  with need, stock, order status - "have we ordered everything?".
'  First click creates the sheet, every later click only refreshes it.
'  Your entries in "Status" (My status, Owner, My comment) are kept.
'
'  Windchill sits behind Microsoft sign-in. Update BOM opens Windchill in
'  the Windows browser engine (hidden; shown only if you have to sign in)
'  and runs the read-only CAD structure query inside that page.
'
'  Nothing is ever written to Windchill, SAP or the order list.
' =====================================================================

' ----- sheet names -----
Private Const SH_START As String = "Start"
Private Const SH_STATUS As String = "Status"
Private Const SH_BOM As String = "BOM"
Private Const SH_OL As String = "Order list"
Private Const SH_SAP As String = "SAP"
Private Const SH_REQ As String = "Auslagerung"
Private Const SH_CSP As String = "CSP Hall"

' ----- Windchill -----
Private Const WC_DEFAULT As String = "https://lp-global-plm.abb.com/Windchill"

' ----- SAP -----
Private Const SAP_TCODE As String = "LX02"
Private Const SAP_LAYOUT_ROW As Long = 8          ' line of your layout in "Choose layout" (as recorded)
Private Const SAP_EXPORT As String = "SHIELD_SAP_EXPORT.XLSX"
Private Const DUMMY_DEFAULT As String = "GCE1093120P0111"

' ----- SCM order list: sheet with "Part ID" in column C, data from row 7 (1-based columns) -----
Private Const OL_FIRST_ROW As Long = 7
Private Const OL_PART As Long = 3
Private Const OL_REV As Long = 4
Private Const OL_DESC As Long = 5
Private Const OL_COMMENT As Long = 26
Private Const OL_SAPID As Long = 28
Private Const OL_DECMS As Long = 31
Private Const OL_SUPPLIER As Long = 38
Private Const OL_ORDERNO As Long = 47
Private Const OL_ORDERQTY As Long = 49
Private Const OL_EXPECTED As Long = 55
Private Const OL_STATUS As Long = 61
Private Const OL_LASTCOL As Long = 61

' ----- layout -----
Private Const BOM_HDR As Long = 3                 ' header row on sheet BOM (data below)
Private Const ST_HDR As Long = 3                  ' header row on sheet Status
Private Const REQ_HDR As Long = 6                 ' header row on sheet Auslagerung
Private Const CSP_HDR As Long = 1                 ' header row on sheet CSP Hall (copy of the CSP Hall file)
Private Const CSP_FILE_DEFAULT As String = "SHIELD_CSP_Hall.xlsx"   ' in the same folder as this tool


' =====================================================================
'  1  UPDATE BOM  (Windchill live)
' =====================================================================
Public Sub UpdateBOM()
    Dim assy As String, rows As Collection, errMsg As String
    assy = CleanNum(Setting("AssyNo"))
    If assy = "" Then
        MsgBox "Please enter the assembly number on the Start sheet first.", vbExclamation
        Exit Sub
    End If
    Set rows = WcReadStructure(assy, errMsg)
    Application.StatusBar = False
    If rows Is Nothing Then
        MsgBox "Could not read the BOM of " & assy & " from Windchill." & vbCrLf & vbCrLf & errMsg & vbCrLf & vbCrLf & _
               "Nothing was changed. Please send a screenshot of this message.", vbExclamation
        Exit Sub
    End If
    WriteBOM rows, "Windchill " & Format$(Now, "dd.mm.yyyy hh:nn")
    RefreshStatus True
    MsgBox "BOM of " & assy & " loaded from Windchill: " & rows.Count & " lines." & vbCrLf & _
           "See sheets ""BOM"" and ""Status"".", vbInformation
End Sub

' ---------------------------------------------------------------------
'  Windchill is behind Microsoft sign-in. Excel therefore opens Windchill
'  in the Windows browser engine (signed in with your account - if a
'  sign-in is needed, the window is shown), and runs the read-only REST
'  query (CAD structure, GetStructure) inside that page.
'  Returns Collection of Array(level, number, version, title, qty per parent, excl) or Nothing.
' ---------------------------------------------------------------------
Private Function WcReadStructure(ByVal assy As String, ByRef errMsg As String) As Collection
    Dim ie As Object, win As Object, doc As Object, base As String, ans As String
    Dim lines() As String, i As Long, f As Variant, rows As Collection, t As Single

    On Error GoTo Fail
    base = Setting("WCBase"): If base = "" Then base = WC_DEFAULT
    If Right$(base, 1) = "/" Then base = Left$(base, Len(base) - 1)

    Application.StatusBar = "Windchill: opening (signed in with your Microsoft account) ..."
    On Error Resume Next
    Set ie = CreateObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")   ' medium mode, keeps the connection
    If ie Is Nothing Then Set ie = CreateObject("InternetExplorer.Application")
    On Error GoTo Fail
    If ie Is Nothing Then Err.Raise vbObjectError + 601, , "The Windows browser engine could not be started."
    ie.Visible = False
    ie.Navigate base & "/app/"

    ' hidden first; if no Windchill page after 20 s, show the window so you can sign in
    Set win = WaitForWindchill(ie, base, 20)
    If win Is Nothing Then
        On Error Resume Next
        ie.Visible = True
        ShowAnyWindchillWindow
        On Error GoTo Fail
        Application.StatusBar = "Windchill: please sign in in the browser window (waiting up to 3 minutes) ..."
        Set win = WaitForWindchill(ie, base, 180)
        If win Is Nothing Then Err.Raise vbObjectError + 602, , "No Windchill page within 3 minutes (sign-in not finished?)."
    End If
    Set doc = win.Document

    Application.StatusBar = "Windchill: reading the structure of " & assy & " ..."
    ans = RunInPage(doc, StructureScript(assy))
    On Error Resume Next
    win.Quit
    If Not ie Is Nothing Then ie.Quit
    On Error GoTo Fail

    If Left$(ans, 2) <> "OK" Then Err.Raise vbObjectError + 603, , "Windchill: " & IIf(ans = "", "no answer from the page", Left$(ans, 400))
    lines = Split(Replace(ans, vbCr, ""), vbLf)
    Set rows = New Collection
    For i = 1 To UBound(lines)
        If lines(i) <> "" Then
            f = Split(lines(i), vbTab)
            If UBound(f) >= 5 Then rows.Add Array(CLng(Val(f(0))), CStr(f(1)), CStr(f(2)), CStr(f(3)), ToNum(f(4)), CStr(f(5)))
        End If
    Next i
    If rows.Count = 0 Then Err.Raise vbObjectError + 604, , "Windchill returned an empty structure."
    Set WcReadStructure = rows
    Exit Function
Fail:
    errMsg = Err.Description
    Set WcReadStructure = Nothing
    On Error Resume Next
    If Not ie Is Nothing Then ie.Quit
End Function

' the read-only query, run inside the signed-in Windchill page (plain JavaScript for the IE engine)
Private Function StructureScript(ByVal assy As String) As String
    Dim j As String
    j = "var B='/Windchill/servlet/odata/v5';var P=(typeof JSON!='undefined')?function(t){return JSON.parse(t);}:function(t){return eval('('+t+')');};" & _
        "function g(u){var x=new XMLHttpRequest();x.open('GET',B+u,false);x.setRequestHeader('Accept','application/json');x.send();" & _
        "if(x.status!=200)throw 'answer '+x.status+' '+x.responseText.substr(0,200);return P(x.responseText);}" & _
        "var v=g('/CADDocumentMgmt/CADDocuments?$filter=Number%20eq%20%27" & assy & "%27&$select=ID,Number,Latest').value||[];" & _
        "var d=null;for(var i=0;i<v.length;i++){if(v[i].Latest){d=v[i];break;}}if(!d&&v.length)d=v[0];" & _
        "if(!d)throw 'no CAD document with the number " & assy & " in Windchill';" & _
        "var n='';try{n=g('/PTC/GetCSRFToken()').NonceValue||'';}catch(e){}" & _
        "var S='CADDocument($select=Number,Name,Version,ComAbbTITLE_LOCAL_LANG_PLT_ATT),CADDocumentUse($select=Quantity,EFBOM,Suppressed)';" & _
        "var x=new XMLHttpRequest();x.open('POST',B+'/CADDocumentMgmt/CADDocuments(%27'+d.ID+'%27)/PTC.CADDocumentMgmt.GetStructure?$expand='+S+',Components($expand='+S+';$levels=max)',false);" & _
        "x.setRequestHeader('Accept','application/json');x.setRequestHeader('Content-Type','application/json');if(n)x.setRequestHeader('CSRF_NONCE',n);x.send('{}');" & _
        "if(x.status!=200)throw 'structure '+x.status+' '+x.responseText.substr(0,300);" & _
        "var s=P(x.responseText);if(s.value)s=s.value[0];var L=[];" & _
        "function w(o,l){var c=o.CADDocument||{},u=o.CADDocumentUse||{};if(l&&u.Suppressed)return;" & _
        "var q=l?(u.Quantity==null?1:u.Quantity):1;var t=(c.ComAbbTITLE_LOCAL_LANG_PLT_ATT||o.CADDocumentName||c.Name||'').replace(/[\t\r\n]/g,' ');" & _
        "L.push([l,o.CADDocumentNumber||c.Number||'',c.Version||'',t,q,u.EFBOM?'yes':''].join('\t'));" & _
        "var k=o.Components||[];for(var i=0;i<k.length;i++)w(k[i],l+1);}" & _
        "w(s,0);R='OK\n'+L.join('\n');"
    StructureScript = j
End Function

' Runs JavaScript inside the page; the script sets R, which is returned ("ERR ..." on failure).
Private Function RunInPage(doc As Object, ByVal js As String) As String
    Static n As Long
    Dim el As Object, sc As Object, id As String, par As Object
    n = n + 1
    id = "shieldres" & n
    On Error Resume Next
    Set par = doc.body
    If par Is Nothing Then Set par = doc.documentElement
    On Error GoTo 0
    Set el = doc.createElement("textarea")              ' fresh helper field (no lookup -> no "null")
    el.id = id
    el.style.display = "none"
    par.appendChild el
    Set sc = doc.createElement("script")
    sc.Text = "(function(){var R='';try{" & js & "}catch(e){R='ERR '+(e.message||e);}document.getElementById('" & id & "').value=R;})();"
    par.appendChild sc
    RunInPage = CStr(el.Value)
End Function

' the browser window that shows a loaded Windchill page (the engine sometimes moves the page to a new window)
Private Function WaitForWindchill(ie As Object, ByVal base As String, ByVal seconds As Long) As Object
    Dim t As Single, w As Object, u As String, hostPart As String
    hostPart = Mid$(base, InStr(base, "://") + 3)
    t = Timer
    Do While Timer - t < seconds
        On Error Resume Next
        u = "": u = ie.LocationURL
        If InStr(1, u, hostPart, vbTextCompare) > 0 Then
            If Not ie.Busy And ie.ReadyState = 4 Then Set WaitForWindchill = ie: Exit Function
        End If
        For Each w In CreateObject("Shell.Application").Windows
            u = "": u = w.LocationURL
            If InStr(1, u, hostPart, vbTextCompare) > 0 Then
                If Not w.Busy And w.ReadyState = 4 Then Set WaitForWindchill = w: Exit Function
            End If
        Next w
        On Error GoTo 0
        Application.Wait Now + TimeSerial(0, 0, 1)
        DoEvents
    Loop
End Function

Private Sub ShowAnyWindchillWindow()
    Dim w As Object
    On Error Resume Next
    For Each w In CreateObject("Shell.Application").Windows
        If InStr(1, w.LocationURL, "login.microsoftonline", vbTextCompare) > 0 Or InStr(1, w.LocationURL, "Windchill", vbTextCompare) > 0 Then w.Visible = True
    Next w
End Sub

Private Sub WriteBOM(rows As Collection, ByVal source As String)
    Dim ws As Worksheet, n As Long, i As Long, it As Variant, out() As Variant
    Dim lev() As Long, qpp() As Double, qpu() As Double, stack() As Long, sp As Long
    Dim isNew As Boolean, r As Long

    n = rows.Count
    ReDim out(1 To n, 1 To 8): ReDim lev(1 To n): ReDim qpp(1 To n): ReDim qpu(1 To n): ReDim stack(1 To n + 1)
    i = 0
    For Each it In rows
        i = i + 1
        lev(i) = it(0): qpp(i) = it(4)
    Next it
    ' quantity per unit = product of the quantities up the tree
    For i = 1 To n
        Do While sp > 0
            If lev(stack(sp)) >= lev(i) Then sp = sp - 1 Else Exit Do
        Loop
        If sp = 0 Then qpu(i) = 1 Else qpu(i) = qpp(i) * qpu(stack(sp))
        sp = sp + 1: stack(sp) = i
    Next i
    i = 0
    For Each it In rows
        i = i + 1
        out(i, 1) = lev(i)
        out(i, 2) = CleanNum(CStr(it(1)))
        out(i, 3) = String$(4 * lev(i), " ") & it(3)
        If i = 1 Then out(i, 4) = "" Else out(i, 4) = qpp(i)
        out(i, 5) = qpu(i)
        out(i, 6) = ""                              ' "Counted as" - filled by RefreshStatus
        out(i, 7) = it(2)
        out(i, 8) = it(5)
    Next it

    Set ws = GetSheet(SH_BOM, isNew)
    Application.ScreenUpdating = False
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 9
    ws.Range("A1").Value = "BOM " & out(1, 2) & " " & out(1, 7) & " – " & Trim$(out(1, 3)) & "   (" & source & ")"
    ws.Range("A1").Font.Size = 12: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(31, 56, 100)
    HeaderRow ws, BOM_HDR, Array("Level", "Part number", "Description (indented by level)", "Qty per parent", "Qty per unit", "Counted as", "Version", "Windchill: excl. from BOM")
    ws.Range(ws.Cells(BOM_HDR + 1, 2), ws.Cells(BOM_HDR + n, 2)).NumberFormat = "@"
    ws.Cells(BOM_HDR + 1, 1).Resize(n, 8).Value = out
    ws.Range(ws.Cells(BOM_HDR + 1, 4), ws.Cells(BOM_HDR + n, 5)).NumberFormat = "0.###"
    With ws.Range(ws.Cells(BOM_HDR + 1, 1), ws.Cells(BOM_HDR + n, 8))
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(217, 217, 217)
    End With
    SetWidths ws, Array(6, 18, 60, 9, 9, 16, 8, 10)
    ws.Range(ws.Cells(BOM_HDR, 1), ws.Cells(BOM_HDR + n, 8)).AutoFilter
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(BOM_HDR + 1, 3).Select
    ActiveWindow.FreezePanes = True
    ThisWorkbook.Worksheets(SH_START).Activate

    SetSetting "BOMInfo", out(1, 2) & " " & out(1, 7) & " · " & n & " lines · " & source
End Sub

' =====================================================================
'  2  UPDATE SAP  (LX02 -> export -> sheet "SAP")
' =====================================================================
Public Sub UpdateSAP()
    Dim sess As Object, folder As String, fullPath As String, pathSet As Boolean, t As Single
    Dim wh As String, pick As Variant

    Set sess = GetSapSession()
    If sess Is Nothing Then Exit Sub
    wh = Setting("SAPWarehouse"): If wh = "" Then wh = "240"

    folder = Environ$("TEMP")
    fullPath = folder & "\" & SAP_EXPORT
    CloseIfOpen SAP_EXPORT
    On Error Resume Next
    If Dir(fullPath) <> "" Then Kill fullPath
    If Dir(fullPath) <> "" Then
        MsgBox "The old SAP export is still open somewhere:" & vbCrLf & fullPath & vbCrLf & "Please close it and try again.", vbExclamation
        Exit Sub
    End If
    On Error GoTo Fail

    Application.StatusBar = "SAP: running " & SAP_TCODE & " ..."
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & SAP_TCODE
    sess.findById("wnd[0]").sendVKey 0
    sess.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = wh
    sess.findById("wnd[0]/tbar[1]/btn[8]").press
    ' choose layout
    sess.findById("wnd[0]/tbar[1]/btn[33]").press
    sess.findById("wnd[1]/usr/lbl[1," & SAP_LAYOUT_ROW & "]").SetFocus
    sess.findById("wnd[1]").sendVKey 2
    ' List > Export > Spreadsheet
    Application.StatusBar = "SAP: exporting ..."
    sess.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").Select
    sess.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error Resume Next
    sess.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder
    sess.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = SAP_EXPORT
    pathSet = (Err.Number = 0)
    Err.Clear
    On Error GoTo Fail
    sess.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error Resume Next
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    sess.findById("wnd[0]").sendVKey 0
    On Error GoTo Fail

    If pathSet Then
        t = Timer
        Do While Dir(fullPath) = "" And Timer - t < 30
            DoEvents
        Loop
        Application.Wait Now + TimeSerial(0, 0, 2)
    End If
    If Not pathSet Or Dir(fullPath) = "" Then
        pick = Application.GetOpenFilename("Excel files (*.xls*),*.xls*", , "Select the file SAP just exported")
        If VarType(pick) = vbBoolean Then Application.StatusBar = False: Exit Sub
        fullPath = CStr(pick)
    End If
    CloseIfOpen Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    LoadSapExport fullPath
    Application.StatusBar = False
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "SAP update stopped: " & Err.Description & vbCrLf & vbCrLf & _
           "Check that SAP is open on the SAP Easy Access screen and try again.", vbExclamation
End Sub

' Alt+F8: load an LX02 export you saved yourself
Public Sub LoadSAPFromFile()
    Dim pick As Variant
    pick = Application.GetOpenFilename("Excel files (*.xls*),*.xls*", , "Select the SAP export (LX02)")
    If VarType(pick) = vbBoolean Then Exit Sub
    CloseIfOpen Mid$(CStr(pick), InStrRev(CStr(pick), "\") + 1)
    LoadSapExport CStr(pick)
End Sub

Private Function GetSapSession() As Object
    Dim sapGui As Object, app As Object
    On Error GoTo NoSAP
    Set sapGui = GetObject("SAPGUI")
    Set app = sapGui.GetScriptingEngine
    Set GetSapSession = app.Children(0).Children(0)
    Exit Function
NoSAP:
    MsgBox "SAP is not open. Please start SAP Logon, log on (SAP Easy Access) and click Update SAP again.", vbExclamation
    Set GetSapSession = Nothing
End Function

Private Sub LoadSapExport(ByVal path As String)
    Dim wb As Workbook, v As Variant, out() As Variant, hdr As Long, r As Long, c As Long, j As Long, n As Long
    Dim col(1 To 8) As Long, names As Variant, alt As Variant, h As String, ws As Worksheet, isNew As Boolean
    Dim nDummy As Long, dummy As String

    Application.ScreenUpdating = False
    Application.StatusBar = "Loading the SAP export ..."
    Set wb = Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True)
    v = wb.Worksheets(1).UsedRange.Value
    wb.Close SaveChanges:=False

    For r = 1 To Application.Min(15, UBound(v, 1))
        For c = 1 To UBound(v, 2)
            If Trim$(SafeStr(v(r, c))) = "Material" Then hdr = r: Exit For
        Next c
        If hdr > 0 Then Exit For
    Next r
    If hdr = 0 Then
        Application.ScreenUpdating = True: Application.StatusBar = False
        MsgBox "No 'Material' column in the SAP export - the SAP sheet was not changed.", vbExclamation
        Exit Sub
    End If
    names = Array("Material", "Storage Location", "Material Description", "Storage Type", "Storage Bin", "Available stock", "GR Date", "Certificate Number")
    alt = Array("Material", "Lagerort", "Materialkurztext", "Lagertyp", "Lagerplatz", "Verf" & ChrW(252) & "gbarer Bestand", "Wareneingangsdatum", "Zeugnis-Nr")
    For j = 1 To 8
        For c = 1 To UBound(v, 2)
            h = Trim$(SafeStr(v(hdr, c)))
            If StrComp(h, names(j - 1), vbTextCompare) = 0 Or StrComp(h, alt(j - 1), vbTextCompare) = 0 Then col(j) = c: Exit For
        Next c
    Next j
    If col(6) = 0 Or col(8) = 0 Then
        Application.ScreenUpdating = True: Application.StatusBar = False
        MsgBox "The export has no 'Available stock' or 'Certificate Number' column." & vbCrLf & _
               "Probably the wrong layout was chosen in SAP - the SAP sheet was not changed.", vbExclamation
        Exit Sub
    End If

    dummy = Setting("DummyMat"): If dummy = "" Then dummy = DUMMY_DEFAULT
    n = UBound(v, 1) - hdr
    ReDim out(1 To n, 1 To 8)
    For r = 1 To n
        For j = 1 To 8
            If col(j) > 0 Then out(r, j) = v(hdr + r, col(j))
        Next j
        out(r, 1) = Trim$(SafeStr(out(r, 1))): out(r, 8) = Trim$(SafeStr(out(r, 8)))
        out(r, 6) = ToNum(out(r, 6))
        If out(r, 1) = dummy Then nDummy = nDummy + 1
    Next r

    Set ws = GetSheet(SH_SAP, isNew)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 9
    HeaderRow ws, 1, Array("Material", "Storage location", "Description", "Storage type", "Storage bin", "Available stock", "GR date", "Certificate no.")
    ws.Range("A2:A" & n + 1).NumberFormat = "@"
    ws.Range("H2:H" & n + 1).NumberFormat = "@"
    ws.Range("G2:G" & n + 1).NumberFormat = "dd.mm.yyyy"
    ws.Range("A2").Resize(n, 8).Value = out
    ws.Range("A1:H" & n + 1).AutoFilter
    SetWidths ws, Array(18, 10, 36, 10, 12, 10, 11, 16)
    ws.Tab.Color = RGB(128, 128, 128)

    SetSetting "SAPInfo", "LX02 warehouse " & IIf(Setting("SAPWarehouse") = "", "240", Setting("SAPWarehouse")) & " · " & n & " rows (" & nDummy & " SHIELD dummy) · " & Format$(Now, "dd.mm.yyyy hh:nn")
    Application.StatusBar = False
    RefreshStatus True
    Application.ScreenUpdating = True
    MsgBox "SAP stock updated: " & n & " rows, " & nDummy & " of them SHIELD dummy stock (" & dummy & ").", vbInformation
End Sub

Private Sub CloseIfOpen(ByVal fname As String)
    Dim i As Long
    For i = Application.Workbooks.Count To 1 Step -1
        If StrComp(Application.Workbooks(i).Name, fname, vbTextCompare) = 0 Then Application.Workbooks(i).Close SaveChanges:=False
    Next i
End Sub

' =====================================================================
'  3  UPDATE ORDER LIST  (SharePoint file - opened READ ONLY, never saved)
' =====================================================================
Public Sub UpdateOrderList()
    Dim path As String, fname As String, wb As Workbook, ws As Worksheet, sh As Worksheet
    Dim openedHere As Boolean, secOld As Long, v As Variant, lastRow As Long, r As Long, msg As String

    path = CleanPath(Setting("OLPath"))
    If path = "" Then
        path = CleanPath(InputBox("Paste the path of the SHIELD order list." & vbCrLf & vbCrLf & _
               "Open the order list in the Excel desktop app > File > Info > Copy path.", "Order list path"))
        If path = "" Then Exit Sub
        SetSetting "OLPath", path
    End If
    If InStr(1, path, "/:x:/", vbTextCompare) > 0 Then
        MsgBox "This is a sharing link - Excel cannot open it directly." & vbCrLf & _
               "Open the order list in the Excel desktop app > File > Info > Copy path, and paste that on the Start sheet.", vbExclamation
        Exit Sub
    End If

    fname = FileNameOf(path)
    Application.StatusBar = "Opening the order list (read only) ..."
    Application.ScreenUpdating = False
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, fname, vbTextCompare) = 0 And wb.Name <> ThisWorkbook.Name Then Exit For
    Next wb
    If wb Is Nothing Then
        secOld = Application.AutomationSecurity
        Application.AutomationSecurity = 3                   ' no macros from that file
        Application.DisplayAlerts = False
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, _
                                IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        msg = Err.Description
        On Error GoTo 0
        Application.DisplayAlerts = True
        Application.AutomationSecurity = secOld
        If wb Is Nothing Then
            Application.ScreenUpdating = True: Application.StatusBar = False
            MsgBox "Could not open the order list:" & vbCrLf & path & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & _
                   "Check the path on the Start sheet (File > Info > Copy path) and that you are signed in to SharePoint.", vbExclamation
            Exit Sub
        End If
        openedHere = True
    End If

    On Error GoTo Fail
    ' the sheet with "Part ID" as header of column C (the file also holds Windchill BOM sheets)
    For Each sh In wb.Worksheets
        For r = 1 To OL_FIRST_ROW - 1
            If LCase$(Trim$(SafeStr(sh.Cells(r, OL_PART).Value))) = "part id" Then Set ws = sh: Exit For
        Next r
        If Not ws Is Nothing Then Exit For
    Next sh
    If ws Is Nothing Then Err.Raise vbObjectError + 301, , "No sheet with 'Part ID' in column C found in " & fname & "."
    lastRow = ws.Cells(ws.Rows.Count, OL_PART).End(xlUp).Row
    If lastRow < OL_FIRST_ROW Then Err.Raise vbObjectError + 302, , "No data below row " & OL_FIRST_ROW & " on sheet '" & ws.Name & "'."
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, OL_LASTCOL)).Value
    msg = fname & " / " & ws.Name
    If openedHere Then wb.Close SaveChanges:=False
    Set wb = Nothing

    WriteOrderList v, msg
    Application.StatusBar = False
    Exit Sub
Fail:
    msg = Err.Description
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Update order list stopped: " & msg, vbExclamation
End Sub

Private Sub WriteOrderList(v As Variant, ByVal source As String)
    Dim d As Object, ord As Collection, r As Long, k As Long, pid As String, cls As String
    Dim a As Variant, out() As Variant, ws As Worksheet, isNew As Boolean, ex As Variant, key As Variant, nLines As Long

    Set d = CreateObject("Scripting.Dictionary")
    Set ord = New Collection
    For r = OL_FIRST_ROW To UBound(v, 1)
        pid = CleanNum(SafeStr(v(r, OL_PART)))
        If pid <> "" Then
            nLines = nLines + 1
            cls = OrderClass(SafeStr(v(r, OL_STATUS)))
            ex = ParseDateLoose(v(r, OL_EXPECTED))
            If Not d.Exists(pid) Then
                ' 0 rev, 1 desc, 2 order status, 3 expected, 4 qty, 5 order no, 6 supplier, 7 SAP ID, 8 DECMS, 9 raw status, 10 comment, 11 lines
                d.Add pid, Array("", "", "", Empty, 0#, "", "", "", Empty, "", "", 0)
                ord.Add pid
            End If
            a = d(pid)
            a(0) = AddUnique(a(0), SafeStr(v(r, OL_REV)))
            If a(1) = "" Then a(1) = Trim$(SafeStr(v(r, OL_DESC)))
            ' several lines per part: On order beats Delivered beats Not ordered
            If cls = "On order" Then
                If a(2) <> "On order" Then a(3) = Empty
                a(2) = "On order": a(3) = MaxV(a(3), ex)
            ElseIf a(2) <> "On order" Then
                If cls = "Delivered" Or a(2) = "" Then
                    If a(2) <> "Delivered" Then a(2) = cls
                End If
                a(3) = MaxV(a(3), ex)
            End If
            If IsNumeric(v(r, OL_ORDERQTY)) And Not IsEmpty(v(r, OL_ORDERQTY)) Then a(4) = a(4) + CDbl(v(r, OL_ORDERQTY))
            a(5) = AddUnique(a(5), SafeStr(v(r, OL_ORDERNO)))
            a(6) = AddUnique(a(6), SafeStr(v(r, OL_SUPPLIER)))
            If a(7) = "" Then a(7) = CleanNum(SafeStr(v(r, OL_SAPID)))
            If IsNumeric(v(r, OL_DECMS)) And Not IsEmpty(v(r, OL_DECMS)) Then a(8) = MaxV(a(8), CDbl(v(r, OL_DECMS)))
            a(9) = AddUnique(a(9), SafeStr(v(r, OL_STATUS)))
            a(10) = AddUnique(a(10), SafeStr(v(r, OL_COMMENT)))
            a(11) = a(11) + 1
            d(pid) = a
        End If
    Next r
    If d.Count = 0 Then Err.Raise vbObjectError + 303, , "No part numbers found in column C of the order list."

    ReDim out(1 To d.Count, 1 To 13)
    For Each key In ord
        k = k + 1
        a = d(key)
        out(k, 1) = key: out(k, 2) = a(0): out(k, 3) = a(1): out(k, 4) = a(2): out(k, 5) = a(3)
        If a(4) <> 0 Then out(k, 6) = a(4)
        out(k, 7) = a(5): out(k, 8) = a(6): out(k, 9) = a(7): out(k, 10) = a(8): out(k, 11) = a(9): out(k, 12) = a(10): out(k, 13) = a(11)
    Next key

    Application.ScreenUpdating = False
    Set ws = GetSheet(SH_OL, isNew)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 9
    HeaderRow ws, 1, Array("Part number", "Rev", "Description", "Order status", "Expected", "Order qty", "Order no.", "Supplier", "SAP ID", "Stock DECMS", "Status (order list)", "Comment (order list)", "Lines")
    ws.Range("A2:A" & k + 1).NumberFormat = "@"
    ws.Range("G2:G" & k + 1).NumberFormat = "@"
    ws.Range("I2:I" & k + 1).NumberFormat = "@"
    ws.Range("E2:E" & k + 1).NumberFormat = "dd.mm.yyyy"
    ws.Range("A2").Resize(k, 13).Value = out
    ws.Range("A1:M" & k + 1).AutoFilter
    SetWidths ws, Array(18, 6, 32, 12, 11, 8, 14, 30, 18, 9, 20, 40, 6)
    ws.Tab.Color = RGB(128, 128, 128)

    SetSetting "OLInfo", source & " · " & d.Count & " parts · " & Format$(Now, "dd.mm.yyyy hh:nn")
    RefreshStatus True
    Application.ScreenUpdating = True
    MsgBox "Order list updated: " & nLines & " lines, " & d.Count & " part numbers." & vbCrLf & _
           "The order list itself was opened read-only and not changed.", vbInformation
End Sub

Private Function OrderClass(ByVal s As String) As String
    s = LCase$(Trim$(Replace(s, Chr(160), " ")))
    If s Like "confirmed*" Or s Like "p.o*" Or s Like "po *" Or s Like "wait for conf*" Or s Like "waiting for conf*" Then
        OrderClass = "On order"
    ElseIf s Like "available*" Or s Like "on site*" Then
        OrderClass = "Delivered"
    Else
        OrderClass = "Not ordered"
    End If
End Function

' =====================================================================
'  STATUS  (rebuilt after every update; your columns are kept)
' =====================================================================
Public Sub RefreshStatus(Optional ByVal quiet As Boolean = False)
    Dim wsB As Worksheet, wsS As Worksheet, wsO As Worksheet, wsP As Worksheet, isNew As Boolean
    Dim units As Double, today As Date, dummy As String
    Dim ol As Object, sapCert As Object, sapMat As Object, keep As Object, items As Object, ord As Collection
    Dim v As Variant, n As Long, i As Long, r As Long, last As Long, lev() As Long, kids() As Boolean, cls() As String
    Dim boughtLev As Long, key As Variant, it As Variant, o As Variant, out() As Variant, k As Long
    Dim need As Double, sapS As Double, lineS As Double, st As String, haveOL As Boolean, haveSAP As Boolean
    Dim cnt As Object, prio() As Long, idx() As Long, j As Long, tmp As Long, s As String
    Dim cMy As Long, cOwn As Long, cCom As Long, cspIn As Object, cspReq As Object, cspKan As Object, wsC As Worksheet, cspQ As Double

    If Not SheetExists(SH_BOM) Then
        If Not quiet Then MsgBox "Click Update BOM first.", vbInformation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set wsB = ThisWorkbook.Worksheets(SH_BOM)
    units = ToNum(Setting("Units")): If units <= 0 Then units = 1
    today = Date
    dummy = Setting("DummyMat"): If dummy = "" Then dummy = DUMMY_DEFAULT

    ' order list: part -> row values
    Set ol = CreateObject("Scripting.Dictionary")
    If SheetExists(SH_OL) Then
        Set wsO = ThisWorkbook.Worksheets(SH_OL)
        last = wsO.Cells(wsO.Rows.Count, 1).End(xlUp).Row
        If last >= 2 Then
            haveOL = True
            v = wsO.Range("A2:M" & last).Value
            For i = 1 To UBound(v, 1)
                s = SafeStr(v(i, 1))
                If s <> "" And Not ol.Exists(s) Then ol.Add s, Array(v(i, 4), v(i, 5), v(i, 7), v(i, 8), SafeStr(v(i, 9)), v(i, 10), v(i, 11), v(i, 12))
            Next i
        End If
    End If
    ' SAP: SHIELD dummy stock by certificate no., other stock by material
    Set sapCert = CreateObject("Scripting.Dictionary")
    Set sapMat = CreateObject("Scripting.Dictionary")
    If SheetExists(SH_SAP) Then
        Set wsP = ThisWorkbook.Worksheets(SH_SAP)
        last = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
        If last >= 2 Then
            haveSAP = True
            v = wsP.Range("A2:H" & last).Value
            For i = 1 To UBound(v, 1)
                s = SafeStr(v(i, 1))
                If s = dummy Then
                    sapCert(SafeStr(v(i, 8))) = sapCert(SafeStr(v(i, 8))) + ToNum(v(i, 6))
                ElseIf s <> "" Then
                    sapMat(s) = sapMat(s) + ToNum(v(i, 6))
                End If
            Next i
        End If
    End If

    ' BOM: which rows count (parts; assemblies ordered complete count as one part)
    last = wsB.Cells(wsB.Rows.Count, 2).End(xlUp).Row
    If last <= BOM_HDR Then Application.ScreenUpdating = True: Exit Sub
    v = wsB.Range(wsB.Cells(BOM_HDR + 1, 1), wsB.Cells(last, 5)).Value
    n = UBound(v, 1)
    ReDim lev(1 To n): ReDim kids(1 To n): ReDim cls(1 To n)
    For i = 1 To n: lev(i) = CLng(ToNum(v(i, 1))): Next i
    For i = 1 To n - 1
        If lev(i + 1) > lev(i) Then kids(i) = True
    Next i
    Set items = CreateObject("Scripting.Dictionary")
    Set ord = New Collection
    boughtLev = -1
    For i = 1 To n
        If boughtLev >= 0 And lev(i) <= boughtLev Then boughtLev = -1
        s = SafeStr(v(i, 2))
        If boughtLev >= 0 Then
            cls(i) = "Included"
        ElseIf i = 1 Then
            cls(i) = "Assembly"
        ElseIf kids(i) Then
            cls(i) = "Assembly"
            If ol.Exists(s) Then
                o = ol(s)
                If o(0) = "On order" Or o(0) = "Delivered" Then cls(i) = "Ordered complete": boughtLev = lev(i)
            End If
        Else
            cls(i) = "Part"
        End If
        If cls(i) = "Part" Or cls(i) = "Ordered complete" Then
            If Not items.Exists(s) Then items.Add s, Array(Trim$(SafeStr(v(i, 3))), 0#): ord.Add s
            it = items(s): it(1) = it(1) + ToNum(v(i, 5)): items(s) = it
        End If
    Next i
    ' "Counted as" back to the BOM sheet
    For i = 1 To n
        r = BOM_HDR + i
        wsB.Cells(r, 6).Value = cls(i)
        Select Case cls(i)
            Case "Assembly"
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Interior.Color = RGB(242, 242, 242)
                wsB.Range(wsB.Cells(r, 2), wsB.Cells(r, 3)).Font.Bold = True
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Font.Color = RGB(0, 0, 0)
            Case "Ordered complete"
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Interior.Color = RGB(221, 235, 247)
                wsB.Range(wsB.Cells(r, 2), wsB.Cells(r, 3)).Font.Bold = True
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Font.Color = RGB(0, 0, 0)
            Case "Included"
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Interior.Pattern = xlNone
                wsB.Range(wsB.Cells(r, 2), wsB.Cells(r, 3)).Font.Bold = False
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Font.Color = RGB(150, 150, 150)
            Case Else
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Interior.Pattern = xlNone
                wsB.Range(wsB.Cells(r, 2), wsB.Cells(r, 3)).Font.Bold = False
                wsB.Range(wsB.Cells(r, 1), wsB.Cells(r, 8)).Font.Color = RGB(0, 0, 0)
        End Select
    Next i

    ' your entries from the old Status sheet: part -> (My status, Owner, My comment, description)
    Set keep = CreateObject("Scripting.Dictionary")
    Set wsS = GetSheet(SH_STATUS, isNew)
    last = wsS.Cells(wsS.Rows.Count, 2).End(xlUp).Row
    If Not isNew And last > ST_HDR Then
        cMy = HeadCol(wsS, ST_HDR, "Manual status"): If cMy = 0 Then cMy = HeadCol(wsS, ST_HDR, "My status")
        cOwn = HeadCol(wsS, ST_HDR, "Owner")
        cCom = HeadCol(wsS, ST_HDR, "Comment"): If cCom = 0 Then cCom = HeadCol(wsS, ST_HDR, "My comment")
        o = wsS.Range(wsS.Cells(ST_HDR + 1, 1), wsS.Cells(last, 20)).Value
        For i = 1 To UBound(o, 1)
            s = SafeStr(o(i, 2))
            If s <> "" And Not keep.Exists(s) And cMy > 0 Then keep.Add s, Array(o(i, cMy), o(i, cOwn), o(i, cCom), o(i, 3))
        Next i
    End If

    ' CSP Hall log: part -> qty in CSP Hall / qty requested
    Set cspIn = CreateObject("Scripting.Dictionary"): Set cspReq = CreateObject("Scripting.Dictionary"): Set cspKan = CreateObject("Scripting.Dictionary")
    If SheetExists(SH_CSP) Then
        Set wsC = ThisWorkbook.Worksheets(SH_CSP)
        last = wsC.Cells(wsC.Rows.Count, 2).End(xlUp).Row
        If last > CSP_HDR Then
            o = wsC.Range(wsC.Cells(CSP_HDR + 1, 1), wsC.Cells(last, 8)).Value
            For i = 1 To UBound(o, 1)
                s = SafeStr(o(i, 2))
                If s <> "" Then
                    If LCase$(SafeStr(o(i, 5))) Like "in csp hall*" Then
                        cspIn(s) = cspIn(s) + ToNum(o(i, 4))
                    ElseIf LCase$(SafeStr(o(i, 5))) Like "kanban*" Then
                        cspKan(s) = cspKan(s) + ToNum(o(i, 4))
                    ElseIf LCase$(SafeStr(o(i, 5))) <> "cancelled" Then
                        cspReq(s) = cspReq(s) + ToNum(o(i, 4))
                    End If
                End If
            Next i
        End If
    End If

    ' build the rows
    k = items.Count
    For Each key In keep.Keys
        If Not items.Exists(key) Then
            o = keep(key)
            If SafeStr(o(0)) <> "" Or SafeStr(o(1)) <> "" Or SafeStr(o(2)) <> "" Then k = k + 1
        End If
    Next key
    ReDim out(1 To Application.Max(k, 1), 1 To 14): ReDim prio(1 To Application.Max(k, 1))
    k = 0
    For Each key In ord
        k = k + 1
        it = items(key)
        need = it(1) * units
        out(k, 2) = key: out(k, 3) = it(0): out(k, 4) = need
        sapS = 0: lineS = 0
        If haveSAP Then
            If sapCert.Exists(Right$(key, 10)) Then sapS = sapCert(Right$(key, 10))
            If sapMat.Exists(key) Then lineS = sapMat(key)
            out(k, 5) = sapS
        End If
        If ol.Exists(key) Then
            o = ol(key)
            If o(4) <> "" And o(4) <> key Then
                If sapMat.Exists(o(4)) Then lineS = lineS + sapMat(o(4))
            End If
            If ToNum(o(5)) > lineS Then lineS = ToNum(o(5))           ' DECMS / Kanban (9999)
            out(k, 8) = o(0): out(k, 9) = o(1)
            out(k, 10) = SafeStr(o(3)) & IIf(SafeStr(o(2)) <> "", " / PO " & SafeStr(o(2)), "")
            out(k, 11) = SafeStr(o(6)) & IIf(SafeStr(o(7)) <> "", " – " & SafeStr(o(7)), "")
        ElseIf haveOL Then
            out(k, 8) = "Not in order list"
        Else
            out(k, 8) = "(order list not loaded)"
        End If
        If haveSAP Or lineS > 0 Then out(k, 6) = lineS
        ' CSP Hall (from the request log)
        cspQ = 0
        If cspIn.Exists(key) Then cspQ = cspIn(key): out(k, 7) = cspQ & " in CSP Hall"
        If cspKan.Exists(key) Then out(k, 7) = Trim$(SafeStr(out(k, 7)) & IIf(SafeStr(out(k, 7)) <> "", ", ", "") & cspKan(key) & " Kanban self-pick")
        If cspReq.Exists(key) Then out(k, 7) = Trim$(SafeStr(out(k, 7)) & IIf(SafeStr(out(k, 7)) <> "", ", ", "") & cspReq(key) & " requested")
        ' status
        If (sapS + lineS >= need Or cspQ >= need) And need > 0 Then
            st = "Available"
        ElseIf out(k, 8) = "Delivered" Then
            st = "Delivered - not in SAP"
        ElseIf out(k, 8) = "On order" Then
            If IsDate(out(k, 9)) Then
                If CDate(out(k, 9)) < today Then st = "Overdue" Else st = "On order"
            Else
                st = "On order"
            End If
        ElseIf out(k, 8) = "Not ordered" Then
            st = "Not ordered"
        ElseIf out(k, 8) = "Not in order list" Then
            st = "Not in order list"
        Else
            st = "Missing"
        End If
        If keep.Exists(key) Then
            o = keep(key)
            out(k, 12) = o(0): out(k, 13) = o(1): out(k, 14) = o(2)
            If SafeStr(out(k, 12)) <> "" Then st = SafeStr(out(k, 12))
        End If
        out(k, 1) = st
        prio(k) = StatusPrio(st)
    Next key
    For Each key In keep.Keys                                ' gone from the BOM but you wrote something
        If Not items.Exists(key) Then
            o = keep(key)
            If SafeStr(o(0)) <> "" Or SafeStr(o(1)) <> "" Or SafeStr(o(2)) <> "" Then
                k = k + 1
                out(k, 1) = "Removed from BOM": out(k, 2) = key: out(k, 3) = o(3): out(k, 4) = 0
                out(k, 12) = o(0): out(k, 13) = o(1): out(k, 14) = o(2)
                prio(k) = StatusPrio("Removed from BOM")
            End If
        End If
    Next key

    ' sort: most urgent first, BOM order inside each status
    ReDim idx(1 To Application.Max(k, 1))
    For i = 1 To k: idx(i) = i: Next i
    For i = 2 To k
        tmp = idx(i): j = i - 1
        Do While j >= 1
            If prio(idx(j)) > prio(tmp) Then idx(j + 1) = idx(j): j = j - 1 Else Exit Do
        Loop
        idx(j + 1) = tmp
    Next i
    Dim sorted() As Variant
    Dim colMap As Variant
    colMap = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 14)     ' supplier/PO and order list comment are not shown
    ReDim sorted(1 To Application.Max(k, 1), 1 To 12)
    For i = 1 To k
        For j = 1 To 12: sorted(i, j) = out(idx(i), colMap(j - 1)): Next j
    Next i

    ' write the Status sheet
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False
    wsS.Cells.Clear
    wsS.Cells.Font.Name = "Arial": wsS.Cells.Font.Size = 9
    wsS.Range("A1").Value = "Status " & SafeStr(wsB.Cells(BOM_HDR + 1, 2).Value) & " – " & Trim$(SafeStr(wsB.Cells(BOM_HDR + 1, 3).Value)) & _
                            " · " & units & " unit(s) · " & Format$(Now, "dd.mm.yyyy hh:nn")
    wsS.Range("A1").Font.Size = 12: wsS.Range("A1").Font.Bold = True: wsS.Range("A1").Font.Color = RGB(31, 56, 100)
    wsS.Range("A2").Value = "Yellow columns are yours (kept on every update). ""Manual status"" overrides the status, e.g. Resolved or Not needed."
    wsS.Range("A2").Font.Color = RGB(128, 128, 128): wsS.Range("A2").Font.Italic = True
    HeaderRow wsS, ST_HDR, Array("Status", "Part number", "Description", "Need", "SAP stock (SHIELD)", "Line / Kanban stock", "CSP Hall", _
                                 "Order status", "Expected", "Manual status", "Owner", "Comment")
    For j = 10 To 12: wsS.Cells(ST_HDR, j).Interior.Color = RGB(127, 96, 0): Next j
    If k > 0 Then
        wsS.Range(wsS.Cells(ST_HDR + 1, 2), wsS.Cells(ST_HDR + k, 2)).NumberFormat = "@"
        wsS.Range(wsS.Cells(ST_HDR + 1, 9), wsS.Cells(ST_HDR + k, 9)).NumberFormat = "dd.mm.yyyy"
        wsS.Range(wsS.Cells(ST_HDR + 1, 4), wsS.Cells(ST_HDR + k, 6)).NumberFormat = "0.##"
        wsS.Cells(ST_HDR + 1, 1).Resize(k, 12).Value = sorted
        With wsS.Range(wsS.Cells(ST_HDR + 1, 1), wsS.Cells(ST_HDR + k, 12))
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(217, 217, 217)
            .VerticalAlignment = xlTop
        End With
        wsS.Range(wsS.Cells(ST_HDR + 1, 10), wsS.Cells(ST_HDR + k + 50, 12)).Interior.Color = RGB(255, 249, 229)
        wsS.Range(wsS.Cells(ST_HDR + 1, 1), wsS.Cells(ST_HDR + k, 1)).Font.Bold = True
        For i = 1 To k
            wsS.Cells(ST_HDR + i, 1).Interior.Color = StatusColor(SafeStr(sorted(i, 1)))
        Next i
        With wsS.Range(wsS.Cells(ST_HDR + 1, 10), wsS.Cells(ST_HDR + k + 50, 10)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                 Formula1:=Join(Array("Resolved", "Not needed", "Available"), Application.International(xlListSeparator))
        End With
    End If
    SetWidths wsS, Array(20, 18, 34, 7, 10, 10, 20, 15, 11, 14, 10, 40)
    wsS.Range(wsS.Cells(ST_HDR, 1), wsS.Cells(ST_HDR + Application.Max(k, 1), 12)).AutoFilter
    wsS.Activate
    ActiveWindow.FreezePanes = False
    wsS.Cells(ST_HDR + 1, 3).Select
    ActiveWindow.FreezePanes = True

    ' counts on the Start sheet
    Set cnt = CreateObject("Scripting.Dictionary")
    For i = 1 To k
        cnt(SafeStr(sorted(i, 1))) = cnt(SafeStr(sorted(i, 1))) + 1
    Next i
    WriteCounts cnt, k
    ThisWorkbook.Worksheets(SH_START).Activate
    Application.ScreenUpdating = True
End Sub

Private Function StatusPrio(ByVal st As String) As Long
    Select Case st
        Case "Overdue": StatusPrio = 1
        Case "Not in order list": StatusPrio = 2
        Case "Not ordered", "Missing": StatusPrio = 3
        Case "Delivered - not in SAP": StatusPrio = 4
        Case "On order": StatusPrio = 5
        Case "Available": StatusPrio = 6
        Case "Removed from BOM": StatusPrio = 8
        Case Else: StatusPrio = 7                          ' Resolved, Not needed ...
    End Select
End Function

Private Function StatusColor(ByVal st As String) As Long
    Select Case st
        Case "Available": StatusColor = RGB(198, 239, 206)
        Case "On order": StatusColor = RGB(255, 230, 153)
        Case "Overdue": StatusColor = RGB(244, 177, 131)
        Case "Delivered - not in SAP": StatusColor = RGB(221, 235, 247)
        Case "Not ordered", "Not in order list", "Missing": StatusColor = RGB(248, 105, 107)
        Case "Removed from BOM": StatusColor = RGB(217, 217, 217)
        Case Else: StatusColor = RGB(189, 215, 238)
    End Select
End Function

Private Sub WriteCounts(cnt As Object, ByVal total As Long)
    Dim ws As Worksheet, top As Range, list As Variant, other As Variant, i As Long, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_START)
    Set top = ws.Range("Counts")
    ws.Range(top, top.Offset(10, 1)).ClearContents
    ws.Range(top, top.Offset(10, 0)).Interior.Pattern = xlNone
    list = Array("Available", "On order", "Delivered - not in SAP", "Overdue", "Not ordered", "Not in order list", "Missing", "Resolved", "Not needed", "Removed from BOM")
    For i = LBound(list) To UBound(list)
        If cnt.Exists(list(i)) Then
            top.Offset(r, 0).Value = list(i): top.Offset(r, 1).Value = cnt(list(i))
            top.Offset(r, 0).Interior.Color = StatusColor(CStr(list(i)))
            cnt.Remove list(i)
            r = r + 1
        End If
    Next i
    For Each other In cnt.Keys                                 ' your own statuses
        top.Offset(r, 0).Value = other: top.Offset(r, 1).Value = cnt(other)
        top.Offset(r, 0).Interior.Color = StatusColor(CStr(other))
        r = r + 1
        If r > 9 Then Exit For
    Next other
    top.Offset(r, 0).Value = "Total parts": top.Offset(r, 1).Value = total
    top.Offset(r, 0).Font.Bold = True: top.Offset(r, 1).Font.Bold = True
End Sub

' =====================================================================
'  4  WAREHOUSE REQUEST (Auslagerung -> CSP Hall) + email draft to Samet
'  Sheet "Auslagerung": parts in SHIELD warehouse stock that are needed
'  for this step and not requested yet. Tick x, adjust qty, click
'  "Create email draft": Outlook (classic) opens a draft - you send it.
'  Requests are logged in sheet "CSP Hall" (set "In CSP Hall" when there).
' =====================================================================
Public Sub WarehouseRequest()
    Dim ws As Worksheet, wsB As Worksheet, wsP As Worksheet, isNew As Boolean
    Dim subA As String, unitsStep As Double, mailTo As String, mailCc As String, intro As String
    Dim v As Variant, n As Long, i As Long, r As Long, iStart As Long, iEnd As Long, baseQ As Double, last As Long
    Dim need As Object, ord As Collection, desc As Object, stock As Object, bins As Object, asked As Object, prevSel As Object
    Dim key As Variant, s As String, cert As String, remaining As Double, q As Double, out() As Variant, k As Long, dummy As String
    Dim o As Variant, lev As Long

    If Not SheetExists(SH_BOM) Then MsgBox "Click 1 Update BOM first.", vbInformation: Exit Sub
    If Not SheetExists(SH_SAP) Then MsgBox "Click 2 Update SAP first - the request needs the SHIELD warehouse stock.", vbInformation: Exit Sub
    RefreshStatus True                                     ' makes sure "Counted as" in BOM is current
    Application.ScreenUpdating = False
    dummy = Setting("DummyMat"): If dummy = "" Then dummy = DUMMY_DEFAULT

    ' settings + your previous ticks from the existing sheet
    Set prevSel = CreateObject("Scripting.Dictionary")
    Set ws = GetSheet(SH_REQ, isNew)
    If isNew Then
        unitsStep = ToNum(Setting("Units")): If unitsStep <= 0 Then unitsStep = 1
        intro = "kannst du bitte folgende SHIELD-Teile (Material " & dummy & ") auslagern und in die CSP-Halle bringen?"
    Else
        subA = CleanNum(SafeStr(ws.Range("C2").Value))
        unitsStep = ToNum(ws.Range("C3").Value): If unitsStep <= 0 Then unitsStep = 1
        mailTo = SafeStr(ws.Range("E2").Value): mailCc = SafeStr(ws.Range("E3").Value): intro = SafeStr(ws.Range("E4").Value)
        last = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        If last > REQ_HDR Then
            o = ws.Range(ws.Cells(REQ_HDR + 1, 1), ws.Cells(last, 4)).Value
            For i = 1 To UBound(o, 1)
                If SafeStr(o(i, 4)) <> "" Then prevSel(SafeStr(o(i, 4))) = Array(o(i, 1), o(i, 2))
            Next i
        End If
    End If

    ' BOM part of this step (whole BOM or one sub-assembly)
    Set wsB = ThisWorkbook.Worksheets(SH_BOM)
    last = wsB.Cells(wsB.Rows.Count, 2).End(xlUp).Row
    v = wsB.Range(wsB.Cells(BOM_HDR + 1, 1), wsB.Cells(last, 6)).Value
    n = UBound(v, 1)
    iStart = 1: iEnd = n: baseQ = 1
    If subA <> "" Then
        iStart = 0
        For i = 1 To n
            If CleanNum(SafeStr(v(i, 2))) = subA Then iStart = i: Exit For
        Next i
        If iStart = 0 Then
            Application.ScreenUpdating = True
            MsgBox "Sub-assembly " & subA & " is not in the BOM. Clear cell C2 for the whole BOM.", vbExclamation
            ws.Activate: Exit Sub
        End If
        lev = CLng(ToNum(v(iStart, 1))): baseQ = ToNum(v(iStart, 5)): If baseQ = 0 Then baseQ = 1
        iEnd = iStart
        Do While iEnd < n
            If CLng(ToNum(v(iEnd + 1, 1))) <= lev Then Exit Do
            iEnd = iEnd + 1
        Loop
    End If
    Set need = CreateObject("Scripting.Dictionary"): Set desc = CreateObject("Scripting.Dictionary"): Set ord = New Collection
    For i = iStart To iEnd
        s = SafeStr(v(i, 6))
        If s = "Part" Or s = "Ordered complete" Then
            key = SafeStr(v(i, 2))
            If Not need.Exists(key) Then need(key) = 0#: desc(key) = Trim$(SafeStr(v(i, 3))): ord.Add key
            need(key) = need(key) + ToNum(v(i, 5)) / baseQ * unitsStep
        End If
    Next i

    ' SHIELD warehouse stock (dummy material) by certificate no. + storage bins
    Set stock = CreateObject("Scripting.Dictionary"): Set bins = CreateObject("Scripting.Dictionary")
    Set wsP = ThisWorkbook.Worksheets(SH_SAP)
    last = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then
        o = wsP.Range("A2:H" & last).Value
        For i = 1 To UBound(o, 1)
            If SafeStr(o(i, 1)) = dummy And ToNum(o(i, 6)) > 0 Then
                cert = SafeStr(o(i, 8))
                stock(cert) = stock(cert) + ToNum(o(i, 6))
                bins(cert) = AddUnique(SafeStr(bins(cert)), SafeStr(o(i, 5)))
            End If
        Next i
    End If
    ' already requested (CSP Hall log, all but cancelled)
    Set asked = CreateObject("Scripting.Dictionary")
    If SheetExists(SH_CSP) Then
        With ThisWorkbook.Worksheets(SH_CSP)
            last = .Cells(.Rows.Count, 2).End(xlUp).Row
            If last > CSP_HDR Then
                o = .Range(.Cells(CSP_HDR + 1, 1), .Cells(last, 8)).Value
                For i = 1 To UBound(o, 1)
                    If SafeStr(o(i, 2)) <> "" And LCase$(SafeStr(o(i, 5))) <> "cancelled" Then asked(SafeStr(o(i, 2))) = asked(SafeStr(o(i, 2))) + ToNum(o(i, 4))
                Next i
            End If
        End With
    End If

    ' rows: in SHIELD stock, still needed, not requested yet
    ReDim out(1 To Application.Max(ord.Count, 1), 1 To 11)
    For Each key In ord
        cert = Right$(key, 10)
        If stock.Exists(cert) Then
            remaining = need(key) - ToNum(asked(key))
            If remaining > 0.0001 Then
                k = k + 1
                q = Application.WorksheetFunction.RoundUp(remaining, 0)
                If q > stock(cert) Then q = stock(cert)
                out(k, 1) = "": out(k, 2) = "": out(k, 3) = q        ' you mark the parts yourself
                If prevSel.Exists(key) Then
                    o = prevSel(key)
                    out(k, 1) = o(0): out(k, 2) = o(1)
                End If
                out(k, 4) = key: out(k, 5) = desc(key): out(k, 6) = need(key): out(k, 7) = stock(cert)
                out(k, 8) = ToNum(asked(key)): out(k, 9) = bins(cert): out(k, 10) = cert
                If stock(cert) < remaining Then out(k, 11) = "only " & stock(cert) & " in SHIELD stock"
            End If
        End If
    Next key

    ' write the sheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    DeleteButtons ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 9
    ws.Range("A1").Value = "Warehouse request (Auslagerung -> CSP Hall) – " & SafeStr(wsB.Cells(BOM_HDR + 1, 2).Value)
    ws.Range("A1").Font.Size = 12: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(31, 56, 100)
    ws.Range("C2").NumberFormat = "@"
    ws.Range("A2").Value = "Sub-assembly (empty = whole BOM)": ws.Range("C2").Value = subA
    ws.Range("A3").Value = "Units for this step": ws.Range("C3").Value = unitsStep
    ws.Range("D2").Value = "Email to (Samet)": ws.Range("E2").Value = mailTo
    ws.Range("D3").Value = "CC": ws.Range("E3").Value = mailCc
    ws.Range("D4").Value = "Text after ""Hallo Samet,""": ws.Range("E4").Value = intro
    ws.Range("A2:A3,D2:D4").Font.Bold = True
    ws.Range("D2:D4").HorizontalAlignment = xlRight
    ws.Range("C2:C3,E2:E4").Interior.Color = RGB(255, 249, 229)
    ws.Range("A5").Value = "Mark the parts you need: x in column A and/or your quantity in column B (x alone = suggested qty). Then ""Create email draft"". After changing C2/C3 click ""Refresh list""."
    ws.Range("A5").Font.Italic = True: ws.Range("A5").Font.Color = RGB(128, 128, 128)
    HeaderRow ws, REQ_HDR, Array("Request (x)", "Your qty", "Suggested qty", "Part number", "Description", "Need (this step)", _
                                 "SHIELD stock (warehouse)", "Already requested", "Storage bin", "Certificate no.", "Note")
    ws.Cells(REQ_HDR, 1).Interior.Color = RGB(127, 96, 0): ws.Cells(REQ_HDR, 2).Interior.Color = RGB(127, 96, 0)
    If k > 0 Then
        ws.Range(ws.Cells(REQ_HDR + 1, 4), ws.Cells(REQ_HDR + k, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(REQ_HDR + 1, 10), ws.Cells(REQ_HDR + k, 10)).NumberFormat = "@"
        ws.Cells(REQ_HDR + 1, 1).Resize(k, 11).Value = out
        ws.Range(ws.Cells(REQ_HDR + 1, 6), ws.Cells(REQ_HDR + k, 8)).NumberFormat = "0.##"
        ws.Range(ws.Cells(REQ_HDR + 1, 1), ws.Cells(REQ_HDR + k, 2)).Interior.Color = RGB(255, 249, 229)
        ws.Range(ws.Cells(REQ_HDR + 1, 1), ws.Cells(REQ_HDR + k, 1)).HorizontalAlignment = xlCenter
        With ws.Range(ws.Cells(REQ_HDR + 1, 1), ws.Cells(REQ_HDR + k, 11))
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(217, 217, 217)
        End With
        ws.Range(ws.Cells(REQ_HDR, 1), ws.Cells(REQ_HDR + k, 11)).AutoFilter
    Else
        ws.Cells(REQ_HDR + 1, 4).Value = "Nothing to request: the needed parts are not in SHIELD stock, or already requested (see sheet ""CSP Hall"")."
    End If
    SetWidths ws, Array(10, 9, 9, 18, 34, 9, 10, 10, 16, 14, 24)
    AddButton ws, ws.Range("I2"), "Refresh list", "WarehouseRequest", RGB(118, 118, 118)
    AddButton ws, ws.Range("K2"), "Create email draft", "CreateSametDraft", RGB(31, 56, 100)
    AddButton ws, ws.Range("K4"), "Mark as sent", "MarkRequestSent", RGB(118, 118, 118)
    ws.Activate
    Application.ScreenUpdating = True
    If k = 0 Then MsgBox "Nothing to request right now - see the note on sheet ""Auslagerung"".", vbInformation
End Sub

Public Sub CreateSametDraft()
    Dim ws As Worksheet, last As Long, i As Long, n As Long, o As Variant, sel As Collection, it As Variant
    Dim mailTo As String, mailCc As String, intro As String, subj As String, html As String, rows As String, warn As String
    Dim olApp As Object, m As Object, body As String, p As Long, assy As String, dummy As String

    Set sel = MarkedRequestRows()
    If sel Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_REQ)
    dummy = Setting("DummyMat"): If dummy = "" Then dummy = DUMMY_DEFAULT

    ' checks before drafting
    For Each it In sel                                   ' it = Array(part, desc, qty, stock, bin, cert)
        If it(2) > it(3) Then warn = warn & vbCrLf & "• " & it(0) & ": " & it(2) & " requested, only " & it(3) & " in SHIELD stock"
        If it(4) = "" Then warn = warn & vbCrLf & "• " & it(0) & ": no storage bin in SAP"
    Next it
    If warn <> "" Then
        If MsgBox("Please check:" & warn & vbCrLf & vbCrLf & "Create the draft anyway?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    For Each it In sel
        n = n + 1
        rows = rows & "<tr>" & Td(CStr(n)) & Td(dummy) & Td(CStr(it(5))) & Td(CStr(it(0))) & Td(CStr(it(1))) & _
               Td(CStr(it(4))) & Td(CStr(it(2))) & "</tr>"
    Next it
    mailTo = Trim$(SafeStr(ws.Range("E2").Value))
    If mailTo = "" Then
        mailTo = Trim$(InputBox("E-mail address of Samet (saved in cell E2):", "To"))
        If mailTo <> "" Then ws.Range("E2").Value = mailTo
    End If
    mailCc = Trim$(SafeStr(ws.Range("E3").Value))
    intro = Trim$(SafeStr(ws.Range("E4").Value))
    assy = SafeStr(ThisWorkbook.Worksheets(SH_BOM).Cells(BOM_HDR + 1, 2).Value)
    subj = "SHIELD – Auslagerung in die CSP-Halle – " & assy & IIf(SafeStr(ws.Range("C2").Value) <> "", " / " & SafeStr(ws.Range("C2").Value), "") & _
           " – " & Format$(Date, "dd.mm.yyyy")
    html = "<div style=""font-family:Arial;font-size:10pt"">" & _
           "<p>Hallo Samet,</p><p>" & HtmlEsc(intro) & "</p>" & _
           "<table style=""border-collapse:collapse;font-family:Arial;font-size:9pt"" border=""1"" cellpadding=""4"">" & _
           "<tr style=""background:#1F3864;color:#FFFFFF;font-weight:bold"">" & Th("Pos.") & Th("Material") & Th("Zeugnis-Nr.") & _
           Th("Teilenummer") & Th("Bezeichnung") & Th("Lagerplatz") & Th("Menge") & "</tr>" & rows & "</table>" & _
           "<p>Vielen Dank und viele Grüße</p></div>"

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo Fail
    If olApp Is Nothing Then Err.Raise vbObjectError + 701, , "Outlook (classic) could not be started."
    Set m = olApp.CreateItem(0)
    m.To = mailTo
    m.CC = mailCc
    m.Subject = subj
    m.Display                                             ' shows the draft and adds your signature - it is NOT sent
    body = m.HTMLBody
    p = InStr(1, body, "<body", vbTextCompare)
    If p > 0 Then p = InStr(p, body, ">")
    If p > 0 Then m.HTMLBody = Left$(body, p) & html & Mid$(body, p + 1) Else m.HTMLBody = html & body
    ws.Range("A4").Value = subj                           ' remembered for "Mark as sent"
    ws.Range("A4").Font.Color = RGB(128, 128, 128)

    If MsgBox("The draft to Samet is open in Outlook (" & n & " parts)." & vbCrLf & _
              "Read it and click Send in Outlook." & vbCrLf & vbCrLf & _
              "Have you sent the email?" & vbCrLf & vbCrLf & _
              "Yes = log the " & n & " parts as ""Requested"" in sheet ""CSP Hall""." & vbCrLf & _
              "No = nothing is changed (if you send it later, click ""Mark as sent"").", _
              vbYesNo + vbQuestion, "Email to Samet") = vbYes Then
        LogRequest sel, subj
    End If
    Exit Sub
Fail:
    MsgBox "Email draft stopped: " & Err.Description, vbExclamation
End Sub

' button on sheet Auslagerung: log the marked parts if you sent the mail later
Public Sub MarkRequestSent()
    Dim sel As Collection, subj As String
    Set sel = MarkedRequestRows()
    If sel Is Nothing Then Exit Sub
    subj = SafeStr(ThisWorkbook.Worksheets(SH_REQ).Range("A4").Value)
    If subj = "" Then subj = "Email to Samet " & Format$(Date, "dd.mm.yyyy")
    If MsgBox("Log the " & sel.Count & " marked parts as ""Requested"" (email sent)?", vbYesNo + vbQuestion) = vbYes Then LogRequest sel, subj
End Sub

' marked rows: x in column A and/or a quantity in column B (x alone = suggested qty)
Private Function MarkedRequestRows() As Collection
    Dim ws As Worksheet, last As Long, i As Long, o As Variant, q As Double, c As New Collection
    If Not SheetExists(SH_REQ) Then MsgBox "Click 5 Warehouse request first.", vbInformation: Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_REQ)
    last = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    If last > REQ_HDR Then
        o = ws.Range(ws.Cells(REQ_HDR + 1, 1), ws.Cells(last, 11)).Value
        For i = 1 To UBound(o, 1)
            If SafeStr(o(i, 10)) <> "" Then
                q = ToNum(o(i, 2))
                If q <= 0 And IsTicked(o(i, 1)) Then q = ToNum(o(i, 3))
                If q > 0 Then c.Add Array(SafeStr(o(i, 4)), SafeStr(o(i, 5)), q, ToNum(o(i, 7)), SafeStr(o(i, 9)), SafeStr(o(i, 10)))
            End If
        Next i
    End If
    If c.Count = 0 Then
        MsgBox "No part marked. Put an x in column A and/or your quantity in column B.", vbInformation
        Exit Function
    End If
    Set MarkedRequestRows = c
End Function

Private Sub ClearRequestMarks()
    Dim ws As Worksheet, last As Long
    Set ws = ThisWorkbook.Worksheets(SH_REQ)
    last = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    If last > REQ_HDR Then ws.Range(ws.Cells(REQ_HDR + 1, 1), ws.Cells(last, 2)).ClearContents
    ws.Range("A4").ClearContents
End Sub

Private Function IsTicked(ByVal x As Variant) As Boolean
    Dim s As String
    s = LCase$(Trim$(SafeStr(x)))
    IsTicked = (s = "x" Or s = "yes" Or s = "ja" Or s = "1")
End Function

Private Function Td(ByVal s As String) As String
    Td = "<td>" & HtmlEsc(s) & "</td>"
End Function

Private Function Th(ByVal s As String) As String
    Th = "<th style=""text-align:left"">" & HtmlEsc(s) & "</th>"
End Function

Private Function HtmlEsc(ByVal s As String) As String
    HtmlEsc = Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

' column number of a header text in a header row (0 if missing)
Private Function HeadCol(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    Dim c As Long
    For c = 1 To 30
        If StrComp(Trim$(SafeStr(ws.Cells(r, c).Value)), txt, vbTextCompare) = 0 Then HeadCol = c: Exit Function
    Next c
End Function

Private Sub AddButton(ws As Worksheet, at As Range, ByVal cap As String, ByVal mac As String, ByVal col As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, at.Left, at.Top, 150, 24)
    shp.Name = "btn_" & mac
    shp.Line.Visible = msoFalse
    shp.Fill.ForeColor.RGB = col
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = cap
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Name = "Arial": .TextRange.Font.Size = 10: .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    shp.OnAction = mac
    shp.Placement = xlFreeFloating
End Sub

Private Sub DeleteButtons(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name Like "btn_*" Then ws.Shapes(i).Delete
    Next i
End Sub

' =====================================================================
'  CSP HALL FILE  (SHIELD_CSP_Hall.xlsx in the same folder as this tool)
'  Master list of material requested from / brought to the CSP Hall.
'  "Update CSP Hall" copies it into sheet "CSP Hall" and refreshes Status.
'  Requests to Samet are appended to the file (status "Requested").
'  Columns: Request date | Part number | Description | Qty | Status |
'           Requested as | Storage bin | Note
' =====================================================================
Public Sub UpdateCSPHall()
    UpdateCSPHallInternal False
End Sub

Private Function UpdateCSPHallInternal(ByVal quiet As Boolean) As Boolean
    Dim path As String, wb As Workbook, openedHere As Boolean, v As Variant, last As Long, ws As Worksheet, isNew As Boolean
    Dim n As Long, msg As String, src As Worksheet

    path = CspFilePath()
    Set wb = OpenCspFile(path, True, openedHere, msg)
    If wb Is Nothing Then
        If quiet Then Exit Function
        If msg = "missing" Then
            If MsgBox("The CSP Hall file does not exist yet:" & vbCrLf & path & vbCrLf & vbCrLf & "Create an empty one there?", vbYesNo + vbQuestion) = vbYes Then
                CreateCspFile path
                Set wb = OpenCspFile(path, True, openedHere, msg)
            End If
        Else
            MsgBox "Could not open the CSP Hall file:" & vbCrLf & path & vbCrLf & vbCrLf & msg, vbExclamation
        End If
        If wb Is Nothing Then Exit Function
    End If
    Set src = wb.Worksheets(1)
    last = src.Cells(src.Rows.Count, 2).End(xlUp).Row
    If last < 1 Then last = 1
    v = src.Range(src.Cells(1, 1), src.Cells(last, 8)).Value
    If openedHere Then wb.Close SaveChanges:=False

    Application.ScreenUpdating = False
    Set ws = GetSheet(SH_CSP, isNew)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 9
    HeaderRow ws, CSP_HDR, Array("Request date", "Part number", "Description", "Qty", "Status", "Requested as", "Storage bin", "Note")
    If last > 1 Then
        ws.Range("B2:B" & last).NumberFormat = "@"
        ws.Range("A2").Resize(last - 1, 8).Value = SubArray(v, 2, last)
        ws.Range("A2:A" & last).NumberFormat = "dd.mm.yyyy"
        ws.Range("A1:H" & last).AutoFilter
        n = last - 1
    End If
    SetWidths ws, Array(11, 18, 36, 7, 16, 18, 16, 40)
    ws.Tab.Color = RGB(128, 128, 128)
    SetSetting "CSPInfo", FileNameOf(path) & " · " & n & " lines · " & Format$(Now, "dd.mm.yyyy hh:nn")
    RefreshStatus True
    Application.ScreenUpdating = True
    If Not quiet Then MsgBox "CSP Hall updated: " & n & " lines from " & FileNameOf(path) & "." & vbCrLf & _
                             "To change a status (e.g. ""In CSP Hall""), edit the CSP Hall file and click Update CSP Hall again.", vbInformation
    UpdateCSPHallInternal = True
End Function

' Alt+F8: opens the CSP Hall file for editing
Public Sub OpenCSPHallFile()
    Dim path As String, wb As Workbook, openedHere As Boolean, msg As String
    path = CspFilePath()
    Set wb = OpenCspFile(path, False, openedHere, msg)
    If wb Is Nothing Then
        If msg = "missing" Then CreateCspFile path: Set wb = OpenCspFile(path, False, openedHere, msg)
    End If
    If Not wb Is Nothing Then wb.Activate Else MsgBox "Could not open " & path & vbCrLf & msg, vbExclamation
End Sub

Private Function CspFilePath() As String
    Dim nm As String, base As String, sep As String
    nm = Setting("CSPFile"): If nm = "" Then nm = CSP_FILE_DEFAULT
    If InStr(nm, ":") > 0 Or Left$(nm, 2) = "\\" Then CspFilePath = nm: Exit Function    ' full path given
    base = ThisWorkbook.Path
    If InStr(base, "://") > 0 Then sep = "/" Else sep = "\"
    CspFilePath = base & sep & nm
End Function

' returns the CSP Hall workbook (an already open one is used); msg = "missing" if the file does not exist
Private Function OpenCspFile(ByVal path As String, ByVal readOnly As Boolean, ByRef openedHere As Boolean, ByRef msg As String) As Workbook
    Dim wb As Workbook, fname As String
    openedHere = False: msg = ""
    fname = FileNameOf(path)
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, fname, vbTextCompare) = 0 Then Set OpenCspFile = wb: Exit Function
    Next wb
    If InStr(path, "://") = 0 Then
        If Dir(path) = "" Then msg = "missing": Exit Function
    End If
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=readOnly, AddToMru:=False)
    If wb Is Nothing Then
        msg = Err.Description
        If msg = "" Then msg = "missing"
    End If
    On Error GoTo 0
    If Not wb Is Nothing Then
        If readOnly Then wb.Windows(1).Visible = False
        openedHere = True
    End If
    Set OpenCspFile = wb
End Function

Private Sub CreateCspFile(ByVal path As String)
    Dim wb As Workbook, ws As Worksheet
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "CSP Hall"
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 9
    HeaderRow ws, 1, Array("Request date", "Part number", "Description", "Qty", "Status", "Requested as", "Storage bin", "Note")
    ws.Cells(1, 5).Interior.Color = RGB(127, 96, 0): ws.Cells(1, 8).Interior.Color = RGB(127, 96, 0)
    SetWidths ws, Array(11, 18, 36, 7, 16, 18, 16, 40)
    CspValidation ws, 2
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

Private Sub CspValidation(ws As Worksheet, ByVal lastRow As Long)
    With ws.Range(ws.Cells(2, 5), ws.Cells(lastRow + 500, 5)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:=Join(Array("In CSP Hall", "Kanban self-pick", "Requested", "Cancelled"), Application.International(xlListSeparator))
    End With
End Sub

' rows r1..r2 of a 2D array as a new 2D array (1-based)
Private Function SubArray(v As Variant, ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To r2 - r1 + 1, 1 To UBound(v, 2))
    For i = r1 To r2
        For j = 1 To UBound(v, 2): out(i - r1 + 1, j) = v(i, j): Next j
    Next i
    SubArray = out
End Function

' append the sent request to the CSP Hall file as "Requested", then refresh everything
Private Sub LogRequest(sel As Collection, ByVal subj As String)
    Dim path As String, wb As Workbook, openedHere As Boolean, msg As String, ws As Worksheet, r As Long, it As Variant, dummy As String
    path = CspFilePath()
    Set wb = OpenCspFile(path, False, openedHere, msg)
    If wb Is Nothing And msg = "missing" Then CreateCspFile path: Set wb = OpenCspFile(path, False, openedHere, msg)
    If wb Is Nothing Then
        MsgBox "Could not open the CSP Hall file to log the request:" & vbCrLf & path & vbCrLf & msg, vbExclamation
        Exit Sub
    End If
    If wb.ReadOnly Then
        MsgBox "The CSP Hall file is open read-only (maybe by someone else) - the request was NOT logged:" & vbCrLf & path, vbExclamation
        If openedHere Then wb.Close SaveChanges:=False
        Exit Sub
    End If
    dummy = Setting("DummyMat"): If dummy = "" Then dummy = DUMMY_DEFAULT
    Set ws = wb.Worksheets(1)
    r = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For Each it In sel                                  ' it = Array(part, desc, qty, stock, bin, cert)
        r = r + 1
        ws.Cells(r, 1).Value = Date: ws.Cells(r, 1).NumberFormat = "dd.mm.yyyy"
        ws.Cells(r, 2).NumberFormat = "@": ws.Cells(r, 2).Value = it(0)
        ws.Cells(r, 3).Value = it(1)
        ws.Cells(r, 4).Value = it(2)
        ws.Cells(r, 5).Value = "Requested"
        ws.Cells(r, 6).NumberFormat = "@": ws.Cells(r, 6).Value = dummy & " / " & it(5)
        ws.Cells(r, 7).Value = it(4)
        ws.Cells(r, 8).Value = subj
    Next it
    ws.Range(ws.Cells(2, 1), ws.Cells(r, 8)).Font.Name = "Arial": ws.Range(ws.Cells(2, 1), ws.Cells(r, 8)).Font.Size = 9
    CspValidation ws, r
    wb.Save
    If openedHere Then wb.Close SaveChanges:=False
    UpdateCSPHallInternal True
    ClearRequestMarks
    WarehouseRequest
    MsgBox sel.Count & " parts logged as ""Requested"" in " & FileNameOf(path) & "." & vbCrLf & _
           "When Samet confirms, set them to ""In CSP Hall"" in that file and click Update CSP Hall.", vbInformation
End Sub

' =====================================================================
'  BUTTONS  (run once: Alt+F8 > CreateButtons)
' =====================================================================
Public Sub CreateButtons()
    Dim ws As Worksheet, caps As Variant, macs As Variant, cols As Variant, i As Long, at As Range
    Set ws = ThisWorkbook.Worksheets(SH_START)
    DeleteButtons ws
    caps = Array("1   Update BOM", "2   Update SAP", "3   Update order list", "4   Update CSP Hall", "5   Warehouse request", "Refresh status")
    macs = Array("UpdateBOM", "UpdateSAP", "UpdateOrderList", "UpdateCSPHall", "WarehouseRequest", "RefreshStatusButton")
    For i = 0 To 5
        Set at = Nothing
        On Error Resume Next
        Set at = ws.Range("BtnBOM").Offset(2 * i, 0)
        On Error GoTo 0
        If at Is Nothing Then Set at = ws.Range("E7").Offset(2 * i, 0)
        AddButton ws, at, CStr(caps(i)), CStr(macs(i)), IIf(i = 5, RGB(118, 118, 118), RGB(31, 56, 100))
        ws.Shapes("btn_" & macs(i)).Width = 170: ws.Shapes("btn_" & macs(i)).Height = 26
    Next i
End Sub

Public Sub RefreshStatusButton()
    RefreshStatus False
    If SheetExists(SH_STATUS) Then MsgBox "Status refreshed (units to build: " & Setting("Units") & ").", vbInformation
End Sub

' =====================================================================
'  HELPERS
' =====================================================================
Private Function Setting(ByVal nm As String) As String
    On Error Resume Next
    Setting = Trim$(SafeStr(ThisWorkbook.Names(nm).RefersToRange.Value))
End Function

Private Sub SetSetting(ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    SheetExists = Not ws Is Nothing
End Function

' existing sheet, or a new one in the order Start - Status - BOM - Order list - SAP
Private Function GetSheet(ByVal nm As String, ByRef isNew As Boolean) As Worksheet
    Dim order As Variant, i As Long, after As Worksheet
    isNew = False
    If SheetExists(nm) Then Set GetSheet = ThisWorkbook.Worksheets(nm): Exit Function
    order = Array(SH_START, SH_STATUS, SH_REQ, SH_CSP, SH_BOM, SH_OL, SH_SAP)
    Set after = ThisWorkbook.Worksheets(1)
    For i = LBound(order) To UBound(order)
        If order(i) = nm Then Exit For
        If SheetExists(CStr(order(i))) Then Set after = ThisWorkbook.Worksheets(CStr(order(i)))
    Next i
    Set GetSheet = ThisWorkbook.Worksheets.Add(After:=after)
    GetSheet.Name = nm
    ActiveWindow.DisplayGridlines = False
    isNew = True
End Function

Private Sub HeaderRow(ws As Worksheet, ByVal r As Long, heads As Variant)
    Dim j As Long
    For j = 0 To UBound(heads)
        With ws.Cells(r, j + 1)
            .Value = heads(j)
            .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 56, 100)
            .WrapText = True: .VerticalAlignment = xlCenter
        End With
    Next j
    ws.Rows(r).RowHeight = 26
End Sub

Private Sub SetWidths(ws As Worksheet, w As Variant)
    Dim j As Long
    For j = 0 To UBound(w): ws.Columns(j + 1).ColumnWidth = w(j): Next j
End Sub

' "3XAA105832A0001.SLDASM" / "_POS010" / "_GEN" / "(01)" / "_002" -> "3XAA105832A0001"
Private Function CleanNum(ByVal s As String) As String
    Static re As Object
    Dim i As Long, ext As String, before As String
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "(_POS\d*|_GEN|_\d{3}|\s*\(\d+\))$"
    End If
    s = UCase$(Trim$(Replace(Replace(s, Chr(160), " "), vbTab, " ")))
    i = InStrRev(s, ".")
    If i > 0 Then
        ext = Mid$(s, i + 1)
        If ext = "SLDPRT" Or ext = "SLDASM" Or ext = "SLDDRW" Or ext = "PRT" Or ext = "ASM" Then s = Trim$(Left$(s, i - 1))
    End If
    Do
        before = s
        s = Trim$(re.Replace(s, ""))
    Loop While s <> before And s <> ""
    i = InStr(s, " ")
    If i > 0 Then s = Left$(s, i - 1)
    CleanNum = s
End Function

Private Function CleanPath(ByVal s As String) As String
    s = Trim$(Replace(s, """", ""))
    If InStr(s, "?") > 0 Then s = Left$(s, InStr(s, "?") - 1)
    s = Replace(s, "/:x:/r/", "/", , , vbTextCompare)
    CleanPath = s
End Function

Private Function FileNameOf(ByVal p As String) As String
    FileNameOf = Replace(Mid$(Replace(p, "/", "\"), InStrRev(Replace(p, "/", "\"), "\") + 1), "%20", " ")
End Function

Private Function ToNum(ByVal x As Variant) As Double
    Dim s As String
    If IsError(x) Or IsEmpty(x) Or IsNull(x) Then Exit Function
    If IsNumeric(x) And VarType(x) <> vbString Then ToNum = CDbl(x): Exit Function
    s = Replace(Trim$(CStr(x)), " ", "")
    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then s = Replace(s, ",", ".")
    ToNum = Val(s)
End Function

' real dates, "12.10.2026", "2026-10-12", "KW 43", "CW43/2026" (-> Friday of that week)
Private Function ParseDateLoose(ByVal x As Variant) As Variant
    Dim s As String, re As Object, m As Object, y As Long, w As Long, jan4 As Date
    ParseDateLoose = Empty
    If IsError(x) Or IsEmpty(x) Or IsNull(x) Then Exit Function
    If VarType(x) = vbDate Then ParseDateLoose = CDate(x): Exit Function
    If IsNumeric(x) Then
        If CDbl(x) > 30000 And CDbl(x) < 80000 Then ParseDateLoose = CDate(CDbl(x))
        Exit Function
    End If
    s = Trim$(CStr(x))
    If s = "" Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(\d{1,2})\.(\d{1,2})\.(\d{2,4})"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        y = CLng(m.SubMatches(2)): If y < 100 Then y = y + 2000
        ParseDateLoose = DateSerial(y, CLng(m.SubMatches(1)), CLng(m.SubMatches(0))): Exit Function
    End If
    re.Pattern = "^(\d{4})-(\d{1,2})-(\d{1,2})"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ParseDateLoose = DateSerial(CLng(m.SubMatches(0)), CLng(m.SubMatches(1)), CLng(m.SubMatches(2))): Exit Function
    End If
    re.Pattern = "(KW|CW|WK)\s*(\d{1,2})(\s*[/.\-]\s*(\d{2,4}))?"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        w = CLng(m.SubMatches(1))
        If Len(m.SubMatches(3)) > 0 Then
            y = CLng(m.SubMatches(3)): If y < 100 Then y = y + 2000
        Else
            y = Year(Date)
            If w < DatePart("ww", Date, vbMonday, vbFirstFourDays) - 26 Then y = y + 1
        End If
        jan4 = DateSerial(y, 1, 4)
        ParseDateLoose = jan4 - (Weekday(jan4, vbMonday) - 1) + (w - 1) * 7 + 4
    End If
End Function

Private Function AddUnique(ByVal list As String, ByVal s As String) As String
    s = Trim$(Replace(s, Chr(160), " "))
    If s = "" Then
        AddUnique = list
    ElseIf list = "" Then
        AddUnique = s
    ElseIf InStr(1, " / " & list & " / ", " / " & s & " / ", vbTextCompare) = 0 Then
        AddUnique = list & " / " & s
    Else
        AddUnique = list
    End If
    If Len(AddUnique) > 250 Then AddUnique = Left$(AddUnique, 250)
End Function

Private Function MaxV(ByVal a As Variant, ByVal b As Variant) As Variant
    If IsEmpty(b) Then
        MaxV = a
    ElseIf IsEmpty(a) Then
        MaxV = b
    ElseIf b > a Then
        MaxV = b
    Else
        MaxV = a
    End If
End Function

Private Function SafeStr(ByVal x As Variant) As String
    If IsError(x) Or IsNull(x) Or IsEmpty(x) Then SafeStr = "" Else SafeStr = CStr(x)
End Function
