VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ENTRY_ELEMENT_SPLIT As String = " "
Const ENTRY_GROUP_SPLIT As String = ","
Const ENTRY_PLACEHOLDER As String = "%"

Const DEFAULT_MODIFIEDBY As String = "2.6r3Import"

Const CONSOLE_USER = "<console>"

Private Enum enuDatabaseOperation
    dbopAdd
    dbopRemove
    dbopOverwrite
    dbopNothing
End Enum

Private m_Entries       As Collection       ' of type clsDBEntryObj
Private m_StoragePath   As String           ' Location where the database is stored.
Private m_Loaded        As Boolean          ' TRUE when the database has been loaded.

Public Property Get Entries() As Collection
    Set Entries = m_Entries
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_Loaded
End Property

Public Property Get FilePath() As String
    FilePath = m_StoragePath
End Property

Public Property Let FilePath(ByVal sValue As String)
    m_StoragePath = sValue
End Property


Public Sub Class_Initialize()
    Set m_Entries = New Collection
    m_StoragePath = vbNullString
    m_Loaded = False
End Sub

Public Sub Class_Terminate()
    Set m_Entries = Nothing
    m_StoragePath = vbNullString
    m_Loaded = False
End Sub


' Reads the user database from storage
Public Sub Load(ByVal sPath As String)
    Dim i, j    As Integer
    Dim cRaw    As Collection
    Dim dbObj   As clsDBEntryObj
    Dim bSave   As Boolean          ' TRUE if the database needs to be saved (after conversion)
    
    ' Load the users file into a list
    m_StoragePath = sPath
    Set cRaw = ListFileLoad(m_StoragePath)
    
    ' Clear the database
    Set m_Entries = New Collection
    
    ' Process each entry
    For i = 1 To cRaw.Count
        Set dbObj = ParseEntry(cRaw.Item(i))
        If Not (dbObj Is Nothing) Then
            ' Make sure the user does not already exist
            If Not Database.ContainsEntry(dbObj.Name, dbObj.EntryType) Then
                Call AddEntry(dbObj)
                
                ' Do we need to save? (were new items created?)
                bSave = ((dbObj.CreatedOn = Now) Or bSave)
            End If
        End If
    Next i
    
    ' Add the bot owner
    If Len(Config.BotOwner) > 0 Then
        If Not Database.ContainsEntry(Config.BotOwner) Then
            Set dbObj = CreateNewEntry(Config.BotOwner)
            With dbObj
                .Rank = 200
            End With
            
            Call AddEntry(dbObj)
            bSave = True
        End If
    End If
    
    m_Loaded = True
    
    ' Save the database if needed
    If bSave Then
        Call Save(m_StoragePath)
    End If
End Sub

' Writes the user database to storage
'  If no path is specified, the one it was loaded from will be used.
Public Sub Save(Optional ByVal sPath As String = vbNullString)
    Dim i       As Integer
    Dim cData   As Collection
    
    ' Check if path was omitted (use load path)
    If sPath = vbNullString Then
        sPath = m_StoragePath
    End If
    
    ' If still no path is set, exit.
    If sPath = vbNullString Then
        Exit Sub
    End If
    
    Set cData = New Collection
    
    ' Convert DB entry objects to a list of strings
    For i = 1 To m_Entries.Count
        cData.Add GetEntryData(m_Entries.Item(i))
    Next i
    
    ' Write list to file
    ListFileSave sPath, cData
End Sub

' Adds an entry to the database
Public Sub AddEntry(ByVal oEntry As clsDBEntryObj)
    On Error Resume Next
    m_Entries.Add oEntry, oEntry.ToString()
    On Error GoTo 0
End Sub

' Removes an entry from the database.
Public Sub RemoveEntry(ByVal oEntry As clsDBEntryObj)
    Dim i As Integer
    
    If oEntry Is Nothing Then Exit Sub
    
    On Error Resume Next
    m_Entries.Remove oEntry.ToString()
    
    ' If we are removing a group, we need to unassign it from all of its members.
    If StrComp(oEntry.EntryType, DB_TYPE_GROUP, vbBinaryCompare) = 0 Then
        For i = 1 To m_Entries.Count
            ' Check if the entry is a member
            If m_Entries.Item(i).IsInGroup(oEntry.Name) Then
                m_Entries.Item(i).Groups.Remove oEntry.Name
            End If
        Next i
    End If
    
    On Error GoTo 0
End Sub

' Validates and executes the database modifications requested in a command.
Friend Sub HandleAddCommand(Command As clsCommandObj, _
    ByVal sName As String, Optional ByVal sType As String = vbNullString, _
    Optional ByVal iRank As Integer = -1, Optional ByVal sFlags As String = vbNullString, _
    Optional ByVal sGroup As String = vbNullString, Optional ByVal sBanMessage As String = vbNullString)
    
    Dim doFlagOp        As enuDatabaseOperation     ' The actions being performed to the entry's flags.
    Dim doGroupOp       As enuDatabaseOperation     ' The actions being performed to the entry's groups.
    
    Dim iRequiredRank   As Integer                  ' The calculated rank required to execute the requested modifications.
    Dim oGroupAccess    As udtUserAccess            ' The access available to the group the entry is being added to.
    Dim oUserAccess     As udtUserAccess            ' The access available to the requesting user.
    
    Dim dbCurrentEntry  As clsDBEntryObj            ' The entry being modified or added.
    Dim dbCurrentAccess As udtUserAccess            ' The access available to the entry being modified or added.
    
    Dim sResponse       As String                   ' The response message for a success
    Dim i               As Integer                  ' counter
    Dim sCurrentFlag    As String                   ' flag being processed
    
    ' Check/set defaults
    iRequiredRank = 0
    
    If Len(sName) = 0 Then
        Command.Respond "You must specify a name to add."
        Exit Sub
    End If
    If Len(sType) = 0 Then
        sType = DB_TYPE_USER
    End If
    
    ' Validate rank
    If iRank > 200 Then
        iRank = 200
    ElseIf iRank < -1 Then
        iRank = 0
    End If
    iRequiredRank = iRank
    
    ' Get the access of the user requesting this action.
    If Command.IsLocal Then
        oUserAccess = GetConsoleAccess()
    Else
        oUserAccess = GetUserAccess(Command.Username)
    End If
    
    ' Find out if we already have an entry for this name or we need to make a new one.
    Set dbCurrentEntry = GetEntry(sName, sType)
    If dbCurrentEntry Is Nothing Then
        ' No entry by that name, so make a new one.
        Set dbCurrentEntry = CreateNewEntry(sName, oUserAccess.Username, sType)
    End If
    dbCurrentAccess = GetEntryAccess(dbCurrentEntry)
    
    ' Does the user have access to modify the entry?
    If ((Not Command.IsLocal) And (Not CanUserModifyEntry(oUserAccess.Username, dbCurrentEntry))) Then
        Command.Respond "You do not have sufficient access to modify that entry."
        Exit Sub
    End If
    
    ' Check flags
    '   The first character of the flags can be a + or - to denote a specific action (add/remove)
    '   If no special character is set, the entry's flags will be overridden with the new flags.
    doFlagOp = CheckOperation(sFlags)
    
    If Len(sFlags) > 0 Then
        If Not Config.CaseSensitiveDBFlags Then
            sFlags = UCase(sFlags)
        End If
        
        If Not CheckValidFlags(sFlags) Then
            Command.Respond "The specified flags are invalid."
            Exit Sub
        End If
    End If
    
    ' Check type-specific restrictions
    Select Case sType
        Case DB_TYPE_CLAN
            If Len(sName) < 2 Or Len(sName) > 4 Then
                Command.Respond "The specified clan tag is invalid."
                Exit Sub
            End If
        Case DB_TYPE_GAME
            If StrComp(GetProductInfo(sName).Code, "UNKW", vbBinaryCompare) = 0 Then
                Command.Respond "The specified game is not recognized."
                Exit Sub
            End If
        Case DB_TYPE_GROUP
            If ((Left(sName, 1) = "+") Or (Left(sName, 1) = "-") Or (InStr(1, sName, Space(1)) > 0)) Then
                Command.Respond "The specified group name contains invalid characters."
                Exit Sub
            End If
        Case DB_TYPE_USER
            ' Valid type, but no special handling
        Case Else
            Command.Respond "The specified entry type is invalid."
            Exit Sub
    End Select
    
    ' Check group
    '   Similar to flags, the first character of the group name specifies add/remove.
    doGroupOp = CheckOperation(sGroup)
    
    If doGroupOp <> dbopNothing Then
        ' Ensure that the group exists
        If Not ContainsEntry(sGroup, DB_TYPE_GROUP) Then
            Command.Respond StringFormat("The group ""{0}"" could not be found.", sGroup)
            Exit Sub
        End If
    
        ' Count access required to add to the group.
        oGroupAccess = GetAccess(sGroup, DB_TYPE_GROUP)
        If oGroupAccess.Rank > iRequiredRank Then
            iRequiredRank = oGroupAccess.Rank
        End If
    End If
    
    ' Does the user have the required rank to make these changes?
    If ((Not Command.IsLocal) And (iRequiredRank > 0) And (iRequiredRank >= oUserAccess.Rank)) Then
        Command.Respond "You do not have sufficient access to perform the requested operation(s)."
        Exit Sub
    End If
    
    ' Check for infinite group loops
    If ((doGroupOp = dbopAdd) Or (doGroupOp = dbopOverwrite)) Then
        If StrComp(sName, sGroup, vbTextCompare) = 0 Then
            Command.Respond "You cannot make a group a member of itself."
            Exit Sub
        ElseIf Database.CheckForGroupLoop(sGroup, sName) Then
            Command.Respond StringFormat("You cannot add ""{0}"" to the group ""{1}"" because it is a member of ""{0}"".", sName, sGroup)
            Exit Sub
        End If
    End If
    
    ' If we're removing a group, make sure the entry is a member.
    If doGroupOp = dbopRemove Then
        If Not dbCurrentEntry.IsInGroup(sGroup) Then
            Command.Respond StringFormat("{0} is not a member of the group: {1}", dbCurrentEntry.ToString(), oGroupAccess.Username)
            Exit Sub
        End If
    End If
    
    
    ' All checks and validation are complete. Start modifying the entry and forming a response.
    sResponse = dbCurrentEntry.ToString()
    
    With dbCurrentEntry
        If iRank > -1 Then
            .Rank = iRank
            sResponse = sResponse & " has been given rank " & .Rank & ", "
        End If
    
        If ((Len(sFlags) > 0) And (doFlagOp <> dbopNothing)) Then
            If doFlagOp = dbopOverwrite Then
                .Flags = vbNullString
                doFlagOp = dbopAdd
            End If
        
            ' Handle flag operation
            If doFlagOp = dbopAdd Then
                For i = 1 To Len(sFlags)
                    sCurrentFlag = Mid(sFlags, i, 1)
                    
                    ' Only add the flag if it isn't there already.
                    If Not .HasFlag(sCurrentFlag) Then
                        .Flags = .Flags & sCurrentFlag
                    End If
                Next i
            ElseIf doFlagOp = dbopRemove Then
                For i = 1 To Len(sFlags)
                    sCurrentFlag = Mid(sFlags, i, 1)
                    .Flags = Replace(.Flags, sCurrentFlag, vbNullString)
                Next i
            End If
        
            ' Append to the response
            If Len(.Flags) = 0 Then
                If StrComp(.EntryType, DB_TYPE_USER, vbBinaryCompare) = 0 Then
                    sResponse = sResponse & " has had their flags removed, "
                Else
                    sResponse = sResponse & " has had its flags removed, "
                End If
            Else
                If Right(sResponse, 2) = ", " Then
                    sResponse = sResponse & "flags " & .Flags & ", "
                Else
                    sResponse = sResponse & " has been given flags " & .Flags & ", "
                End If
            End If
        Else
            If Right(sResponse, 2) = ", " Then sResponse = Left(sResponse, Len(sResponse) - 2)
        End If
    
        If ((Len(sGroup) > 0) And (doGroupOp <> dbopNothing)) Then
            If Right(sResponse, 2) = ", " Then
                sResponse = sResponse & "and "
            End If
            sResponse = sResponse & "has been "
            
            If doGroupOp = dbopOverwrite Then
                Call .ClearGroups
                doGroupOp = dbopAdd
            End If
            
            If doGroupOp = dbopAdd Then
                .Groups.Add sGroup, sGroup
                sResponse = sResponse & "made a member of "
            ElseIf doGroupOp = dbopRemove Then
                .Groups.Remove sGroup
                sResponse = sResponse & "removed from "
            End If
            
            sResponse = sResponse & "the group: " & Split(oGroupAccess.Username, Space(1))(0)
        Else
            ' Revise part of the message
            sResponse = Replace(sResponse, ", flags ", " and flags ")
        End If
        
        ' Note: no way to clear ban messages?
        If Len(sBanMessage) > 0 Then
            .BanMessage = sBanMessage
        End If
        
        ' Update metadata
        .ModifiedBy = oUserAccess.Username
        .ModifiedOn = Now()
        
        ' If this entry is new, add it.
        If Not ContainsEntry(.Name, .EntryType) Then
            Call AddEntry(dbCurrentEntry)
        End If
    End With
    
    Call Save
    
    ' Finalize response message
    If Right(sResponse, 2) = ", " Then sResponse = Left(sResponse, Len(sResponse) - 2)
    Command.Respond sResponse & "."
    
    ' Check for changes to channel safety
    Call g_Channel.CheckUsers
    
End Sub


' Returns the first matching entry for the specified name from the database.
'   If a type is supplied then only an entry of the specified type will be returned.
Public Function GetEntry(ByVal sName As String, Optional ByRef sType As Variant = vbNullString) As clsDBEntryObj
    Dim i As Integer
    
    Set GetEntry = Nothing      ' default value
    
    For i = 1 To m_Entries.Count
        With m_Entries.Item(i)
            ' Check entry name
            If StrComp(.Name, sName, vbTextCompare) = 0 Then
            
                ' Check entry type (if null then any type is accepted)
                If ((Len(sType) = 0) Or (StrComp(.EntryType, sType, vbBinaryCompare) = 0)) Then
                    Set GetEntry = m_Entries.Item(i)
                    sType = CStr(GetEntry.EntryType)
                    Exit Function
                End If
            End If
        End With
    Next i
End Function

' Returns a collection of entries for the specified array of names from the database.
'   If a type is supplied then only entries oif the specified type will be returned.
Public Function GetEntries(ByRef aNameList() As Variant, Optional ByVal sType As String = vbNullString) As Collection
    Dim NameCol As Collection
    Dim x       As Integer

    Set NameCol = New Collection

    For x = LBound(aNameList) To UBound(aNameList)
        On Error Resume Next
        NameCol.Add CStr(aNameList(x)), CStr(aNameList(x))
        On Error GoTo 0
    Next x

    Set GetEntries = GetEntriesFromCollection(NameCol, sType)
End Function

' same as the above but from an Entry.Groups collection
Public Function GetEntriesFromCollection(ByVal NameCol As Collection, Optional ByVal sType As String = vbNullString) As Collection
    Dim i As Integer
    Dim j As Integer

    Set GetEntriesFromCollection = New Collection

    For i = 1 To m_Entries.Count
        With m_Entries.Item(i)
            For j = 1 To NameCol.Count
                ' Compare names
                If StrComp(.Name, NameCol.Item(j), vbTextCompare) = 0 Then
                    ' Compare types
                    If ((Len(sType) = 0) Or (StrComp(.EntryType, sType, vbBinaryCompare) = 0)) Then
                        GetEntriesFromCollection.Add m_Entries.Item(i), .ToString()
                    End If
                End If
            Next j
        End With
    Next i
End Function


' Returns TRUE if the database contains an item with the specified name and type.
'   If no name is supplied, any entry with the specified name will match.
Public Function ContainsEntry(ByVal sName As String, Optional ByVal sType As String = vbNullString) As Boolean
    ContainsEntry = Not CBool(GetEntry(sName, sType) Is Nothing)
End Function

' Returns the total level of access available to the entry identified by the given name.
'   If no type is specified then any entry with a matching name will be returned and the type param will be set.
Friend Function GetAccess(ByVal sName As String, Optional ByRef sType As String = vbNullString) As udtUserAccess
    Dim oEntry  As clsDBEntryObj
    Dim bIsUser As Boolean
    
    bIsUser = False
    
    ' Check if we have an entry for the given name and type.
    Set oEntry = GetEntry(sName, sType)
    If (oEntry Is Nothing) Then
        If ((Len(sType) = 0) Or (StrComp(sType, DB_TYPE_USER, vbBinaryCompare) = 0)) Then
            bIsUser = True
        End If
    Else
        ' Is the entry for a user?
        If StrComp(oEntry.EntryType, DB_TYPE_USER, vbBinaryCompare) = 0 Then
            bIsUser = True
        End If
    End If
    
    ' If we think it's a user, check for contextual access (clan, game)
    If bIsUser Then
        sType = DB_TYPE_USER
        GetAccess = GetUserAccess(sName)
    Else
        ' If we have an entry, use that, otherwise return default.
        If oEntry Is Nothing Then
            GetAccess = GetDefaultAccessResponse(sName)
        Else
            sType = oEntry.EntryType
            GetAccess = GetEntryAccess(oEntry)
        End If
    End If
End Function

' Returns TRUE if the specified user/type combo were found.
'  If no type is specified then any entry with a matching name will be returned and the type param will be set.
Friend Function HasAccess(ByVal sName As String, Optional ByRef sType As String = vbNullString) As Boolean
    Dim oAccess As udtUserAccess

    oAccess = GetAccess(sName, sType)

    ' Did we find access?
    If ((oAccess.Rank > 0) Or (Len(oAccess.Flags) > 0)) Then
        HasAccess = True
    End If
End Function

' Returns the level of access available to a given database entry.
'   This level considers all of the groups the entry is assigned to, but not context-based access.
Friend Function GetEntryAccess(ByVal oEntry As clsDBEntryObj) As udtUserAccess
    Dim i       As Integer
    Dim cGroups As Collection

    ' Get default values
    GetEntryAccess = GetDefaultAccessResponse()

    ' Was an entry provided?
    If oEntry Is Nothing Then Exit Function

    With GetEntryAccess
        .Username = oEntry.ToString()

        .Rank = oEntry.Rank
        .Flags = oEntry.Flags

        .BanMessage = oEntry.BanMessage

        Set .Groups = New Collection
        If oEntry.Groups.Count > 0 Then
            Set cGroups = GetEntriesFromCollection(oEntry.Groups, DB_TYPE_GROUP)

            ' Include access for each group this entry is a member of.
            For i = 1 To cGroups.Count
                Call MergeAccess(GetEntryAccess, GetEntryAccess(cGroups.Item(i)))
                .Groups.Add cGroups.Item(i).ToString()
            Next i
        End If
    End With
End Function

' Returns the level of access available to a user.
'   This level includes context-based access available to the user.
Friend Function GetUserAccess(ByVal sUsername As String) As udtUserAccess
    Dim i               As Integer          ' counter
    Dim bMatch          As Boolean          ' True if a match to the current group was found (for dynamic groups)
    Dim oChannelUser    As clsUserObj       ' Holds channel user object representing the search for user.
    Dim cDynamicGroups  As Collection
    
    ' default return value
    GetUserAccess = GetDefaultAccessResponse()
    
    ' Early exit if the database is empty.
    If m_Entries.Count = 0 Then Exit Function
    
    ' Get all of the dynamic, context-based groups.
    Set cDynamicGroups = GetDynamicGroups()
    
    ' Get an entry to base this inquiry off of
    If ContainsEntry(sUsername, DB_TYPE_USER) Then
        GetUserAccess = GetEntryAccess(GetEntry(sUsername, DB_TYPE_USER))
    End If
    
    ' Find the user in the channel
    Set oChannelUser = g_Channel.GetUser(sUsername)
    
    ' Check dynamic groups
    If cDynamicGroups.Count > 0 Then
        For i = 1 To cDynamicGroups.Count
            With cDynamicGroups.Item(i)
                ' Check for pattern matching groups
                If StrComp(.EntryType, DB_TYPE_USER, vbBinaryCompare) = 0 Then
                    ' Try to match the pattern
                    If (LCase(PrepareCheck(sUsername)) Like LCase(PrepareCheck(.Name))) Then
                        bMatch = True
                    End If
                ElseIf Len(oChannelUser.Name) > 0 Then
                
                    ' Check for GAME and CLAN matches
                    If StrComp(.EntryType, DB_TYPE_GAME, vbBinaryCompare) = 0 Then        ' Type: GAME
                        If StrComp(.Name, oChannelUser.Game, vbTextCompare) = 0 Then
                            bMatch = True
                        End If
                    ElseIf StrComp(.EntryType, DB_TYPE_CLAN, vbBinaryCompare) = 0 Then    ' Type: CLAN
                        If StrComp(.Name, oChannelUser.Clan, vbTextCompare) = 0 Then
                            bMatch = True
                        End If
                    End If
                End If
                
                ' Merge with this dynamic group
                If bMatch Then
                    Call MergeAccess(GetUserAccess, Database.GetEntryAccess(cDynamicGroups.Item(i)))
                    GetUserAccess.Groups.Add cDynamicGroups.Item(i).ToString()
                End If
            End With
            
            bMatch = False
        Next i
    End If
    
End Function

' Returns a formatted identifying string for the given access object.
'   Ex: Bob [People (group), BoT (clan)]
Friend Function GetAccessNameString(ByRef oAccess As udtUserAccess) As String
    Dim sReturn As String
    Dim i       As Integer
    
    ' Start with username
    sReturn = oAccess.Username
    
    ' Append groups
    If oAccess.Groups.Count > 0 Then
        If Len(sReturn) > 0 Then
            sReturn = sReturn & Space(1)
        End If
        sReturn = sReturn & "["
        
        For i = 1 To oAccess.Groups.Count
            sReturn = sReturn & oAccess.Groups.Item(i) & ", "
        Next i
        sReturn = Left(sReturn, Len(sReturn) - 2) & "]"
    End If
    
    GetAccessNameString = sReturn
End Function

' Returns a collection of dynamic groups.
'   Dynamic groups are entries that use match characters or context (wildcards, type: GAME/CLAN)
Public Function GetDynamicGroups() As Collection
    Dim i       As Integer
    
    Set GetDynamicGroups = New Collection
    If m_Entries.Count > 0 Then
        For i = 1 To m_Entries.Count
            With m_Entries.Item(i)
                If ((InStr(1, .Name, "*", vbBinaryCompare) > 0) Or _
                    (InStr(1, .Name, "?", vbBinaryCompare) > 0) Or _
                    (.EntryType = DB_TYPE_GAME) Or (.EntryType = DB_TYPE_CLAN)) Then
                    
                    GetDynamicGroups.Add m_Entries.Item(i), .ToString()
                End If
            End With
        Next i
    End If
End Function

' Returns the level of access available to the bot console.
Friend Function GetConsoleAccess() As udtUserAccess
    GetConsoleAccess = GetDefaultAccessResponse(CONSOLE_USER)
    With GetConsoleAccess
        .Rank = 201
        .Flags = "A"
    End With
End Function

' Returns a date representing the last time the database was modified.
'   (based on the LastModified values for individual entries)
Public Function GetLastModified() As Date
    Dim i       As Integer
    
    GetLastModified = "00:00:00 12/30/1899"         ' Set default lowest value
    If m_Entries.Count = 0 Then Exit Function    ' Return default if database is empty.
    
    For i = 1 To m_Entries.Count
        If DateDiff("s", GetLastModified, m_Entries.Item(i).ModifiedOn) > 0 Then
            GetLastModified = m_Entries.Item(i).ModifiedOn
        End If
    Next i
End Function

' Returns TRUE if the specified flag string is valid. (alphabetic)
Public Function CheckValidFlags(ByVal sFlags As String) As Boolean
    Dim iMin    As Integer     ' Minimum valid value
    Dim iMax    As Integer     ' Maximum valid value
    
    iMin = Asc("A")
    iMax = Asc("Z")
    
    Dim v       As Integer     ' value of current character
    Dim i       As Integer     ' counter
    
    sFlags = UCase(sFlags)
    For i = 1 To Len(sFlags)
        v = Asc(Mid(sFlags, i, 1))
        
        If ((v < iMin) Or (v > iMax)) Then
            CheckValidFlags = False
            Exit Function
        End If
    Next i
    CheckValidFlags = True
End Function

' Returns TRUE if there is an infinite loop of group assignments in the specified group.
Public Function CheckForGroupLoop(ByVal sGroupName As String, Optional ByVal sGroupToFind As String = vbNullString) As Boolean
    CheckForGroupLoop = False
    
    Dim oEntry As clsDBEntryObj
    Set oEntry = GetEntry(sGroupName, DB_TYPE_GROUP)
    
    Dim i As Integer
    
    If oEntry Is Nothing Then Exit Function
    If oEntry.Groups.Count = 0 Then Exit Function
    
    ' If no group is specified assume the one being searched.
    If Len(sGroupToFind) = 0 Then
        sGroupToFind = oEntry.Name
    End If
    
    ' Check each group this one is assigned to.
    For i = 1 To oEntry.Groups.Count
        If StrComp(oEntry.Groups.Item(i), sGroupToFind, vbTextCompare) = 0 Then
            CheckForGroupLoop = True
            Exit Function
        Else
            ' Recursively check sub-groups.
            If CheckForGroupLoop(oEntry.Groups.Item(i), sGroupToFind) Then
                CheckForGroupLoop = True
                Exit Function
            End If
        End If
    Next i
End Function

' Returns a new database entry object with the specified values.
'   The object is NOT added to the database automatically.
Public Function CreateNewEntry(ByVal sName As String, Optional ByVal sCreator As String = CONSOLE_USER, Optional ByVal sEntryType As String = DB_TYPE_USER) As clsDBEntryObj
    Set CreateNewEntry = GetDefaultEntry(sName)
    With CreateNewEntry
        .CreatedBy = sCreator
        .ModifiedBy = sCreator
        
        .EntryType = sEntryType
    End With
End Function

' Returns TRUE if the specified username has sufficient access to modify the given database entry.
Public Function CanUserModifyEntry(ByVal sUsername As String, ByVal oEntry As clsDBEntryObj) As Boolean
    Dim oUserAccess     As udtUserAccess        ' Access available to the user making modifications.
    Dim oEntryAccess    As udtUserAccess        ' Access available to the entry being modified.
    
    CanUserModifyEntry = False
    
    oUserAccess = GetUserAccess(sUsername)
    
    ' If the entry is a user, get their contextual access.
    oEntryAccess = GetEntryAccess(oEntry)
    
    ' Check if the entry is locked.
    If oEntry.HasFlag("L") Then
        ' User must have database admin flag (A) or be high enough rank.
        If ((InStr(1, oUserAccess.Flags, "A") = 0) And (oUserAccess.Rank < 100)) Then
            Exit Function
        End If
    End If
    
    ' For unlocked entries, rank is the deciding factor.
    If ((oEntryAccess.Rank = 0) Or (oEntryAccess.Rank < oUserAccess.Rank)) Then
        CanUserModifyEntry = True
        Exit Function
    End If
End Function

' Creates the special config-defined groups.
Public Sub CreateSpecialGroups()
    Dim aGroupNames(2)  As String
    Dim oGroupEntry     As clsDBEntryObj
    Dim i               As Integer
    
    aGroupNames(0) = Config.SafelistGroup
    aGroupNames(1) = Config.ShitlistGroup
    aGroupNames(2) = Config.TagbanGroup
    
    For i = 0 To UBound(aGroupNames)
        ' Check that the group is defined
        If Len(aGroupNames(i)) > 0 Then
            ' Check if the group already exists.
            If Not ContainsEntry(aGroupNames(i), DB_TYPE_GROUP) Then
                Set oGroupEntry = CreateNewEntry(aGroupNames(i), CONSOLE_USER, DB_TYPE_GROUP)
            
                ' Set the group's flags
                Select Case aGroupNames(i)
                    Case Config.SafelistGroup
                        oGroupEntry.Flags = "S"
                    Case Config.ShitlistGroup, Config.TagbanGroup
                        oGroupEntry.Flags = "B"
                End Select
            
                Call AddEntry(oGroupEntry)
            End If
        End If
    Next i
End Sub

' Creates a new database with a copy of all of the entries contained in this database.
'  The file path field is not copied, and would need to be specified manually if saved.
Public Function CreateCopy() As clsDatabase
    Dim i       As Integer
    
    Set CreateCopy = New clsDatabase
    For i = 1 To m_Entries.Count
        Call CreateCopy.AddEntry(m_Entries.Item(i).CreateCopy())
    Next i
End Function

' Merges 2 access objects together, updating BASE to their combined values.
'   Returns TRUE if OTHER had any effect changing the values.
Private Function MergeAccess(ByRef uaBase As udtUserAccess, ByRef uaOther As udtUserAccess) As Boolean
    Dim i       As Integer      ' counter
    Dim sFlag   As String       ' current flag to compare

    MergeAccess = False
    
    ' Highest rank is assigned
    If uaOther.Rank > uaBase.Rank Then
        uaBase.Rank = uaOther.Rank
        MergeAccess = True
    End If
    
    ' Check for new flags
    For i = 1 To Len(uaOther.Flags)
        sFlag = Mid(uaOther.Flags, i, 1)
        
        ' If this flag is not in the base, add it.
        If InStr(1, uaBase.Flags, sFlag, vbBinaryCompare) = 0 Then
            uaBase.Flags = uaBase.Flags & sFlag
            MergeAccess = True
        End If
    Next i
    
    ' Only change the ban message if one isn't already set
    If Len(uaBase.BanMessage) = 0 And Len(uaOther.BanMessage) > 0 Then
        uaBase.BanMessage = uaOther.BanMessage
        MergeAccess = True
    End If
End Function

' Parses a database entry from a line in the user database file.
' Entry format: username access flags createdby createdon modifiedby modifiedon entrytype banmessage
'                  0        1     2       3         4          5          6         7         8
Private Function ParseEntry(ByVal sEntryData As String) As clsDBEntryObj
    On Error GoTo ERROR_HANDLER
    Dim e()         As String   ' entry elements
    Dim i, j        As Integer  ' counters
    Dim aGroups()   As String   ' group list
    
    Set ParseEntry = Nothing
    
    If InStr(1, sEntryData, ENTRY_ELEMENT_SPLIT, vbTextCompare) > 0 Then
        e = Split(sEntryData, ENTRY_ELEMENT_SPLIT, 10)      ' 10 elements
        
        Set ParseEntry = GetDefaultEntry(e(0))
        With ParseEntry
            For i = 1 To UBound(e)
                If e(i) <> ENTRY_PLACEHOLDER Then
                    Select Case i
                        Case 1, 2
                            If StrictIsNumeric(e(i)) Then
                                .Rank = Int(e(i))
                            Else
                                .Flags = e(i)
                            End If
                        Case 3
                            .CreatedBy = e(i)
                        Case 4
                            .CreatedOn = ParseDate(e(i))
                        Case 5
                            .ModifiedBy = e(i)
                        Case 6
                            .ModifiedOn = ParseDate(e(i))
                        Case 7
                            .EntryType = e(i)
                        Case 8
                            aGroups = Split(e(i), ENTRY_GROUP_SPLIT)
                            For j = 0 To UBound(aGroups)
                                If LenB(aGroups(j)) > 0 Then
                                    .Groups.Add aGroups(j), LCase$(aGroups(j))
                                End If
                            Next j
                        Case 9
                            .BanMessage = e(i)
                    End Select
                End If
            Next i
            
            ' Validate some elements
            If Len(.Name) = 0 Then
                Set ParseEntry = Nothing
                Exit Function
            End If
            
            If .Rank > 200 Then
                .Rank = 200
            End If
            
            If Len(.EntryType) = 0 Or .EntryType = ENTRY_PLACEHOLDER Then
                .EntryType = DB_TYPE_USER
            End If
        End With
    End If

    Exit Function
ERROR_HANDLER:
    If Err.Number = 457 Then
        ' entry with two groups with the same name
        Err.Clear
        Resume Next
    End If
End Function

' Returns a string used to store the specified entry in the database file.
Private Function GetEntryData(ByVal oEntry As clsDBEntryObj) As String
    Dim e()         As Variant
    Dim i           As Integer
    Dim DDate       As Date
    Dim aGroups()   As String
    
    With oEntry
        ' Does the entry have group assignments?
        If .Groups.Count > 0 Then
            ' Turn the collection into an array.
            ReDim aGroups(.Groups.Count - 1)
            For i = 1 To .Groups.Count
                aGroups(i - 1) = .Groups.Item(i)
            Next i
        Else
            ReDim aGroups(0)
        End If
            
    
        e = Array(.Name, .Rank, .Flags, .CreatedBy, .CreatedOn, .ModifiedBy, .ModifiedOn, .EntryType, Join(aGroups, ENTRY_GROUP_SPLIT), .BanMessage)
    End With
    
    ' Validate elements
    For i = 0 To UBound(e)
        Select Case i
            Case 4, 6   ' date fields
                DDate = CDate(e(i))
                If DDate > 0 Then
                    e(i) = FormatDate(DDate)
                End If
            Case Else   ' string fields
                If Len(e(i)) = 0 Then
                    e(i) = ENTRY_PLACEHOLDER
                End If
        End Select
    Next i
    
    GetEntryData = Join(e, ENTRY_ELEMENT_SPLIT)
End Function

' Returns a DB entry with default values
Private Function GetDefaultEntry(ByVal sName As String) As clsDBEntryObj
    Set GetDefaultEntry = New clsDBEntryObj
    With GetDefaultEntry
        .Name = sName
        
        .CreatedOn = Now
        .CreatedBy = DEFAULT_MODIFIEDBY
        .ModifiedOn = Now
        .ModifiedBy = DEFAULT_MODIFIEDBY
        
        .Rank = 0
        .Flags = vbNullString
        .EntryType = DB_TYPE_USER
        
        .BanMessage = vbNullString
    End With
End Function

' Returns the default access response. This will indicate that the search returned no results.
Private Function GetDefaultAccessResponse(Optional ByVal sName As String = vbNullString) As udtUserAccess
    With GetDefaultAccessResponse
        .Username = sName
                
        .Flags = vbNullString
        .Rank = 0
        Set .Groups = New Collection

        .BanMessage = vbNullString
    End With
End Function

' Converts a stored date string to a date object
Private Function ParseDate(ByVal sElement As String) As Date
    ParseDate = CDate(Replace(sElement, "_", Space(1)))
End Function

' Converts a date object to a string for storage.
Private Function FormatDate(ByVal DDate As Date) As String
    Dim s As String
    s = Format(DDate, "dd-MM-yyyy_HH:MM:SS")
    FormatDate = Replace(s, Space(1), "_")
End Function

' Determines the operation indicated by the specified string.
'   This parses the +/- from flags and group arguments.
'   Updates the provided string with the indicator removed.
Private Function CheckOperation(ByRef sOpString As String) As enuDatabaseOperation
    If Len(sOpString) = 0 Then
        CheckOperation = dbopNothing
        Exit Function
    End If
    
    Select Case Left(sOpString, 1)
        Case "+"
            CheckOperation = dbopAdd
        Case "-"
            CheckOperation = dbopRemove
        Case Else
            CheckOperation = dbopOverwrite
    End Select
    
    ' Remove the operation indicator
    If Not (CheckOperation = dbopOverwrite) Then
        sOpString = Mid(sOpString, 2)
    End If
End Function


