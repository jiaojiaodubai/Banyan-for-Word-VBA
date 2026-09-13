Option Explicit

' ============================================================================
' Module  : modAutomation
' Purpose : Agent-facing tools for COM-driven automation (Application.Run).
'
'   These entry points finish the Banyan flow unattended: styles come from
'   POST /banyan/style/list and items from Zotero's local API
'   (GET /api/users/0/items...), so the GUI routes /banyan/style and
'   /banyan/citation are never called. Pending fields reuse the modField
'   factories and get their rich text from modRefresh.RefreshInRange. Every
'   public function returns False/Nothing/"" on failure and never shows a
'   dialog.
'
'   Item dict shape built here (a subset of the backend's Banyan item):
'     { id, key, itemType, uri, title, date, year, firstCreator, language,
'       creators, tags, extra }
'   `id` is always 0 (the local API exposes no numeric id); the backend
'   resolves the item by uri when syncItems=True, so a partial snapshot works.
'
' Public API:
'   AutomationStyleList() As Collection
'   AutomationStyleApply(styleId, styleTitle) As Boolean
'   AutomationPreferenceStyleId() As String
'   AutomationItemsSnapshot(limit) As Long
'   AutomationItemKeyAt(index) As String
'   AutomationItemTitleAt(index) As String
'   AutomationItemTypeAt(index) As String
'   AutomationItemGetByKey(itemKey) As Object
'   AutomationCitationSource(itemKeysCsv) As Object
'   AutomationInsertPendingCitation(source) As Boolean
'   AutomationInsertPendingCitationForKey(itemKeysCsv) As Boolean
'   AutomationInsertPendingBibliography() As Boolean
'   AutomationRefresh() As Boolean
' ============================================================================

' Zotero's local API mirrors Web API v3 under /api/users/0; host and port
' follow modHttp (IPv4 loopback, 23119 with dev fallback 23124).
Private Const ZOTERO_API_HOST As String = "127.0.0.1"
Private Const ZOTERO_API_BASE As String = "/api/users/0"
Private Const ZOTERO_URI_ROOT As String = "http://zotero.org/"

Private Const BIBLIOGRAPHY_PLACEHOLDER As String = "{ BIBLIOGRAPHY }"
Private Const PLACEHOLDER_COLOR As String = "#ff0000"

' Cached item snapshot (Banyan item dicts) + key index, filled by
' AutomationItemsSnapshot and consumed by AutomationItemGetByKey.
Private m_snapshot As Collection
Private m_snapshotByKey As Object


' --- Styles -----------------------------------------------------------------

' --- AutomationStyleList - Fetch the Banyan style list (no GUI route). ---
' POST /banyan/style/list -> { ok, data: [ { id, title, citationType, ... } ] }
Public Function AutomationStyleList() As Collection
    On Error GoTo ErrHandler

    Dim respText As String
    respText = HttpPost(HttpBuildUrl("style/list"), "{}")
    If Len(respText) = 0 Then Exit Function

    Dim envelope As Object
    Set envelope = JsonParse(respText)
    If envelope Is Nothing Then Exit Function
    If Not DictKeyBool(envelope, "ok") Then Exit Function
    If Not DictKeyIsCollection(envelope, "data") Then Exit Function

    Set AutomationStyleList = DictKeyObject(envelope, "data")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationStyleList"
    Set AutomationStyleList = Nothing
End Function


' --- AutomationStyleApply - Save a style from the list into BANYAN_PREF. ---
' Resolves by id first, then title; a missing preference is created with
' production defaults (PreferenceSave normalizes the remaining fields).
Public Function AutomationStyleApply(ByVal styleId As String, _
                                     Optional ByVal styleTitle As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim styles As Collection
    Set styles = AutomationStyleList()
    If styles Is Nothing Then Exit Function

    Dim style As Object
    Set style = FindStyleInList(styles, styleId, styleTitle)
    If style Is Nothing Then Exit Function

    Dim pref As Object
    Set pref = PreferenceGet()
    If pref Is Nothing Then Set pref = New Dictionary
    Set pref("style") = GetPrefStyle(DictKeyString(style, "id"), _
                                     DictKeyString(style, "title"), _
                                     DictKeyString(style, "citationType"))

    AutomationStyleApply = PreferenceSave(pref)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationStyleApply"
    AutomationStyleApply = False
End Function


' --- AutomationPreferenceStyleId - Style id stored in the active document. ---
' Returns "" when no preference exists.
Public Function AutomationPreferenceStyleId() As String
    On Error GoTo ErrHandler

    Dim pref As Object
    Set pref = PreferenceGet()
    If pref Is Nothing Then Exit Function

    Dim style As Object
    Set style = DictKeyObject(pref, "style")
    If style Is Nothing Then Exit Function

    AutomationPreferenceStyleId = DictKeyString(style, "id")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationPreferenceStyleId"
    AutomationPreferenceStyleId = ""
End Function


Private Function FindStyleInList(ByVal styles As Collection, _
                                 ByVal styleId As String, _
                                 ByVal styleTitle As String) As Object
    Dim entry As Variant

    If Len(styleId) > 0 Then
        For Each entry In styles
            If DictKeyString(entry, "id") = styleId Then
                Set FindStyleInList = entry
                Exit Function
            End If
        Next entry
    End If

    If Len(styleTitle) > 0 Then
        For Each entry In styles
            If DictKeyString(entry, "title") = styleTitle Then
                Set FindStyleInList = entry
                Exit Function
            End If
        Next entry
    End If
End Function


' --- Items (Zotero local API) ------------------------------------------------

' --- AutomationItemsSnapshot - Cache a local item snapshot. ---
' Uncitable top-level entries (attachments, notes) are skipped. Returns the
' cached item count (0 on failure).
Public Function AutomationItemsSnapshot(Optional ByVal limit As Long = 50) As Long
    On Error GoTo ErrHandler

    Set m_snapshot = New Collection
    Set m_snapshotByKey = New Dictionary
    If limit < 1 Then limit = 50

    Dim respText As String
    respText = HttpGet(ZoteroApiUrl("/items/top?limit=" & CStr(limit) & "&format=json"))
    If Len(respText) = 0 Then Exit Function

    Dim parsed As Object
    Set parsed = JsonParse(respText)
    If parsed Is Nothing Then Exit Function
    If Not DictIsCollection(parsed) Then Exit Function

    Dim record As Variant
    Dim item As Object
    For Each record In parsed
        Set item = BanyanItemFromApiRecord(record)
        If Not item Is Nothing Then
            m_snapshot.Add item
            Dim itemKey As String
            itemKey = DictKeyString(item, "key")
            If Not m_snapshotByKey.Exists(itemKey) Then
                Set m_snapshotByKey(itemKey) = item
            End If
        End If
    Next record

    AutomationItemsSnapshot = m_snapshot.Count
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationItemsSnapshot"
    AutomationItemsSnapshot = 0
End Function


' --- AutomationItemKeyAt / TitleAt / TypeAt - Inspect the snapshot. ---
Public Function AutomationItemKeyAt(ByVal index As Long) As String
    On Error GoTo ErrHandler
    If m_snapshot Is Nothing Then Exit Function
    If index < 1 Or index > m_snapshot.Count Then Exit Function
    AutomationItemKeyAt = DictKeyString(m_snapshot(index), "key")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationItemKeyAt"
    AutomationItemKeyAt = ""
End Function

Public Function AutomationItemTitleAt(ByVal index As Long) As String
    On Error GoTo ErrHandler
    If m_snapshot Is Nothing Then Exit Function
    If index < 1 Or index > m_snapshot.Count Then Exit Function
    AutomationItemTitleAt = DictKeyString(m_snapshot(index), "title")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationItemTitleAt"
    AutomationItemTitleAt = ""
End Function

Public Function AutomationItemTypeAt(ByVal index As Long) As String
    On Error GoTo ErrHandler
    If m_snapshot Is Nothing Then Exit Function
    If index < 1 Or index > m_snapshot.Count Then Exit Function
    AutomationItemTypeAt = DictKeyString(m_snapshot(index), "itemType")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationItemTypeAt"
    AutomationItemTypeAt = ""
End Function


' --- AutomationItemGetByKey - One item by key (cache first, then query). ---
Public Function AutomationItemGetByKey(ByVal itemKey As String) As Object
    On Error GoTo ErrHandler

    itemKey = Trim$(itemKey)
    If Len(itemKey) = 0 Then Exit Function

    If Not m_snapshotByKey Is Nothing Then
        If m_snapshotByKey.Exists(itemKey) Then
            Set AutomationItemGetByKey = m_snapshotByKey(itemKey)
            Exit Function
        End If
    End If

    Dim respText As String
    respText = HttpGet(ZoteroApiUrl("/items?itemKey=" & itemKey & "&format=json"))
    If Len(respText) = 0 Then Exit Function

    Dim parsed As Object
    Set parsed = JsonParse(respText)
    If parsed Is Nothing Then Exit Function
    If Not DictIsCollection(parsed) Then Exit Function
    If parsed.Count = 0 Then Exit Function

    Set AutomationItemGetByKey = BanyanItemFromApiRecord(parsed(1))
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationItemGetByKey"
    Set AutomationItemGetByKey = Nothing
End Function


' --- AutomationCitationSource - Build a CitationSource from item keys. ---
' itemKeysCsv: comma-separated item keys (one cite each). Returns
' { cites: [ { item, params } ], params: {} } or Nothing when none resolve.
Public Function AutomationCitationSource(ByVal itemKeysCsv As String) As Object
    On Error GoTo ErrHandler

    Dim cites As Collection
    Set cites = New Collection

    Dim keys As Variant
    keys = Split(itemKeysCsv, ",")
    Dim i As Long
    Dim itemKey As String
    Dim item As Object
    Dim cite As Object

    For i = LBound(keys) To UBound(keys)
        itemKey = Trim$(CStr(keys(i)))
        If Len(itemKey) > 0 Then
            Set item = AutomationItemGetByKey(itemKey)
            If item Is Nothing Then
                LogAutomation "item not found: " & itemKey
            Else
                Set cite = New Dictionary
                Set cite("item") = item
                Set cite("params") = New Dictionary
                cites.Add cite
            End If
        End If
    Next i

    If cites.Count = 0 Then Exit Function

    Dim source As Object
    Set source = New Dictionary
    Set source("cites") = cites
    Set source("params") = New Dictionary
    Set AutomationCitationSource = source
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationCitationSource"
    Set AutomationCitationSource = Nothing
End Function


' --- Pending inserts (reuse modField field operations) ------------------------

' --- AutomationInsertPendingCitation - Insert a placeholder citation. ---
' Inserts at the document end using the preference's citation type; a later
' refresh replaces the placeholder. Note citations get a Word-numbered
' footnote (no custom reference).
Public Function AutomationInsertPendingCitation(ByVal source As Object) As Boolean
    On Error GoTo ErrHandler

    If source Is Nothing Then Exit Function
    If Not FieldIsCitationSource(source) Then Exit Function

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then Exit Function

    Dim style As Object
    Set style = DictKeyObject(pref, "style")
    If style Is Nothing Then Exit Function

    Dim cursor As Range
    Set cursor = DocumentEndRange()

    Dim data As Object
    Dim noteData As Object
    Dim created As Collection

    Select Case DictKeyString(style, "citationType")
        Case "intext-citation"
            Set data = FieldCreatePlaceholderIntextCitationData(FieldCreateId(), source)
            AutomationInsertPendingCitation = Not (FieldCreateIntextCitationAtRange(cursor, data) Is Nothing)
        Case "note-citation"
            Set noteData = FieldCreatePlaceholderNoteCitationData(FieldCreateId(), source)
            Set created = FieldCreateNoteCitationAtRange(cursor, noteData)
            AutomationInsertPendingCitation = Not (created Is Nothing)
    End Select
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationInsertPendingCitation"
    AutomationInsertPendingCitation = False
End Function


' --- AutomationInsertPendingCitationForKey - COM-friendly variant. ---
' Lets a COM caller pass only strings (no VBA object marshaling).
Public Function AutomationInsertPendingCitationForKey(ByVal itemKeysCsv As String) As Boolean
    On Error GoTo ErrHandler

    Dim source As Object
    Set source = AutomationCitationSource(itemKeysCsv)
    If source Is Nothing Then Exit Function

    AutomationInsertPendingCitationForKey = AutomationInsertPendingCitation(source)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationInsertPendingCitationForKey"
    AutomationInsertPendingCitationForKey = False
End Function


' --- AutomationInsertPendingBibliography - Insert a placeholder block. ---
' One bibliography-title placeholder; the refresh replaces it with the
' returned lines.
Public Function AutomationInsertPendingBibliography() As Boolean
    On Error GoTo ErrHandler

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then Exit Function

    Dim data As Object
    Set data = New Dictionary
    data("id") = FieldCreateId()
    data("type") = "bibliography-title"
    Set data("content") = FieldCreateRichText(BIBLIOGRAPHY_PLACEHOLDER, PLACEHOLDER_COLOR)

    Dim cursor As Range
    Set cursor = DocumentEndRange()

    Dim fld As Field
    Set fld = FieldCreateRawAddinField(cursor, "BANYAN_BIBLIOGRAPHY " & DictKeyString(data, "id"))
    If fld Is Nothing Then Exit Function
    If Not FieldWriteData(fld, data) Then Exit Function

    FieldRenderStyledFieldWithStyle fld, _
                                    DictKeyString(pref, "bibliographyTitleStyle"), _
                                    wdStyleTypeParagraph, _
                                    DictKeyObject(data, "content")
    AutomationInsertPendingBibliography = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationInsertPendingBibliography"
    AutomationInsertPendingBibliography = False
End Function


' --- Refresh ------------------------------------------------------------------

' --- AutomationRefresh - Refresh the whole document with syncItems=True. ---
' Uses the dialog-free core, not RefreshAction.
Public Function AutomationRefresh() As Boolean
    On Error GoTo ErrHandler
    AutomationRefresh = RefreshInRange(ActiveDocument.Content, True)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.AutomationRefresh"
    AutomationRefresh = False
End Function


' --- Zotero API helpers -------------------------------------------------------

Private Function ZoteroApiUrl(ByVal pathAndQuery As String) As String
    ZoteroApiUrl = "http://" & ZOTERO_API_HOST & ":" & HttpGetPort() & ZOTERO_API_BASE & pathAndQuery
End Function


' Build a Banyan item dict from one Web API item record. Field mapping follows
' the backend's toBanyanItem for the fields that drive citation generation.
Private Function BanyanItemFromApiRecord(ByVal record As Variant) As Object
    On Error GoTo ErrHandler

    If Not DictIsObject(record) Then Exit Function

    Dim rec As Object
    Set rec = record

    Dim data As Object
    Set data = DictKeyObject(rec, "data")
    If data Is Nothing Then Exit Function

    Dim itemType As String
    itemType = DictKeyString(data, "itemType")
    If Not IsCitableItemType(itemType) Then Exit Function

    Dim itemKey As String
    itemKey = DictKeyString(data, "key")
    If Len(itemKey) = 0 Then itemKey = DictKeyString(rec, "key")
    If Len(itemKey) = 0 Then Exit Function

    Dim itemUri As String
    itemUri = ItemUriFromApiRecord(rec, itemKey)
    If Len(itemUri) = 0 Then Exit Function

    Dim item As Object
    Set item = New Dictionary
    item("id") = 0
    item("key") = itemKey
    item("itemType") = itemType
    item("uri") = itemUri
    item("title") = DictKeyString(data, "title")
    item("date") = DictKeyString(data, "date")
    item("language") = DictKeyString(data, "language")

    Dim meta As Object
    Set meta = DictKeyObject(rec, "meta")
    If Not meta Is Nothing Then
        item("year") = DictKeyString(meta, "parsedDate")
        item("firstCreator") = DictKeyString(meta, "creatorSummary")
    End If

    If DictKeyIsCollection(data, "creators") Then
        Set item("creators") = DictKeyObject(data, "creators")
    Else
        Set item("creators") = New Collection
    End If
    Set item("tags") = ItemTagsFromApiData(data)
    Set item("extra") = ItemExtraFromApiData(data)

    Set BanyanItemFromApiRecord = item
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.BanyanItemFromApiRecord"
    Set BanyanItemFromApiRecord = Nothing
End Function


' Canonical item URI (http://zotero.org/users|groups/<id>/items/<key>): prefer
' links.self (covers unsynced local libraries), else library type/id.
Private Function ItemUriFromApiRecord(ByVal rec As Object, ByVal itemKey As String) As String
    On Error GoTo ErrHandler

    Dim links As Object
    Set links = DictKeyObject(rec, "links")
    If Not links Is Nothing Then
        Dim selfLink As Object
        Set selfLink = DictKeyObject(links, "self")
        If Not selfLink Is Nothing Then
            Dim href As String
            href = DictKeyString(selfLink, "href")
            Dim apiPos As Long
            apiPos = InStr(1, href, "/api/", vbTextCompare)
            If apiPos > 0 Then
                ItemUriFromApiRecord = ZOTERO_URI_ROOT & Mid$(href, apiPos + 5)
                Exit Function
            End If
        End If
    End If

    Dim library As Object
    Set library = DictKeyObject(rec, "library")
    If library Is Nothing Then Exit Function

    Dim libraryId As String
    libraryId = DictKeyString(library, "id")
    If Len(libraryId) = 0 Then Exit Function

    If DictKeyString(library, "type") = "group" Then
        ItemUriFromApiRecord = ZOTERO_URI_ROOT & "groups/" & libraryId & "/items/" & itemKey
    Else
        ItemUriFromApiRecord = ZOTERO_URI_ROOT & "users/" & libraryId & "/items/" & itemKey
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modAutomation.ItemUriFromApiRecord"
    ItemUriFromApiRecord = ""
End Function


Private Function ItemTagsFromApiData(ByVal data As Object) As Collection
    Dim result As Collection
    Set result = New Collection

    If Not DictKeyIsCollection(data, "tags") Then
        Set ItemTagsFromApiData = result
        Exit Function
    End If

    Dim entry As Variant
    Dim tagText As String
    For Each entry In DictKeyObject(data, "tags")
        tagText = ""
        If DictIsObject(entry) Then
            tagText = DictKeyString(entry, "tag")
            If Len(tagText) = 0 Then tagText = DictKeyString(entry, "name")
        ElseIf DictIsString(entry) Then
            tagText = CStr(entry)
        End If
        If Len(tagText) > 0 Then result.Add tagText
    Next entry

    Set ItemTagsFromApiData = result
End Function


' Parse Zotero's string `extra` field ("key: value" per line) into the map the
' Banyan item exposes. Duplicate keys keep the last value; the backend refetches
' the canonical map when syncItems=True.
Private Function ItemExtraFromApiData(ByVal data As Object) As Object
    Dim result As Object
    Set result = New Dictionary

    Dim text As String
    text = DictKeyString(data, "extra")
    If Len(text) = 0 Then
        Set ItemExtraFromApiData = result
        Exit Function
    End If

    Dim lines As Variant
    lines = Split(Replace(text, vbCr, vbLf), vbLf)

    Dim i As Long
    Dim currentLine As String
    Dim separatorPos As Long
    Dim extraKey As String
    Dim extraValue As String
    For i = LBound(lines) To UBound(lines)
        currentLine = Trim$(CStr(lines(i)))
        If Len(currentLine) > 0 Then
            separatorPos = InStr(1, currentLine, ":", vbBinaryCompare)
            If separatorPos > 1 Then
                extraKey = Trim$(Left$(currentLine, separatorPos - 1))
                extraValue = Trim$(Mid$(currentLine, separatorPos + 1))
                If Len(extraKey) > 0 Then result(extraKey) = extraValue
            End If
        End If
    Next i

    Set ItemExtraFromApiData = result
End Function


Private Function IsCitableItemType(ByVal itemType As String) As Boolean
    If Len(itemType) = 0 Then Exit Function
    Select Case LCase$(itemType)
        Case "attachment", "note", "annotation"
            Exit Function
    End Select
    IsCitableItemType = True
End Function


' --- Document helpers ---------------------------------------------------------

' Collapsed range at the end of the main text. Content.Duplicate + Collapse is
' the safe pattern (doc.Range(End, End) raises 4608 in some Word builds).
Private Function DocumentEndRange() As Range
    Dim target As Range
    Set target = ActiveDocument.Content.Duplicate
    target.Collapse wdCollapseEnd
    Set DocumentEndRange = target
End Function


Private Sub LogAutomation(ByVal message As String)
    Debug.Print "[Banyan][Automation] " & message
End Sub
