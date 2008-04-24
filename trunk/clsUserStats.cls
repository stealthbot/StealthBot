VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsUserStats.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_stat_string         As String
Private m_game                As String
Private m_icon                As String
Private m_spawn               As Boolean
Private m_clan                As String
Private m_level               As Long
Private m_character_name      As String
Private m_character_class_id  As Integer
Private m_character_flags     As Long
Private m_acts_completed      As Long
Private m_wins                As Long
Private m_ladder_rating       As Long
Private m_high_rating         As Long
Private m_ladder_rank         As Long
Private m_strength            As Long
Private m_dexterity           As Long
Private m_vitality            As Long
Private m_gold                As Long
Private m_magic               As Long
Private m_dots                As Long
Private m_expansion           As Boolean
Private m_hardcore            As Boolean
Private m_realm               As String

Public Property Let game(str As String)

    ' ...
    m_game = str

End Property

Public Property Get game() As String

    ' ...
    game = m_game

End Property

Public Property Let Icon(str As String)

    ' ...
    m_icon = str

End Property

Public Property Get Icon() As String

    ' ...
    Icon = m_icon

End Property

Public Property Get Race() As String

    ' ...
    Icon = m_icon

End Property

Public Property Get Tier() As String

    ' ...

End Property

Public Property Get IconName() As String

    ' ...
    If ((game = "WAR3") Or (game = "W3XP")) Then
        ' ...
        Select Case (Mid$(Icon, 2, 1))
            Case "H" ' Human
                ' ...
                If (game = "WAR3") Then
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "footman"
                        Case 3: IconName = "knight"
                        Case 4: IconName = "Archmage"
                        Case 5: IconName = "Medivh"
                    End Select
                Else
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "rifleman"
                        Case 3: IconName = "sorceress"
                        Case 4: IconName = "spellbreaker"
                        Case 5: IconName = "Blood Mage"
                        Case 6: IconName = "Jaina Proudmore"
                    End Select
                End If
            
            Case "N" ' Night Elf
                ' ...
                If (game = "WAR3") Then
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "archer"
                        Case 3: IconName = "druid of the claw"
                        Case 4: IconName = "Priestess of the Moon"
                        Case 5: IconName = "Furion"
                    End Select
                Else
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "huntress"
                        Case 3: IconName = "druid of the talon"
                        Case 4: IconName = "dryad"
                        Case 5: IconName = "Keeper of the Grove"
                        Case 6: IconName = "Maiev"
                    End Select
                End If
                
            Case "U" ' Undead
                ' ...
                If (game = "WAR3") Then
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "ghoul"
                        Case 3: IconName = "abomination"
                        Case 4: IconName = "Lich"
                        Case 5: IconName = "Tichondrius"
                    End Select
                Else
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "crypt fiend"
                        Case 3: IconName = "banshee"
                        Case 4: IconName = "destroyer"
                        Case 5: IconName = "Crypt Lord"
                        Case 6: IconName = "Sylvanas"
                    End Select
                End If
            
            Case "O" ' Orc
                ' ...
                If (game = "WAR3") Then
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "grunt"
                        Case 3: IconName = "tauren"
                        Case 4: IconName = "Far Seer"
                        Case 5: IconName = "Thrall"
                    End Select
                Else
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "headhunter"
                        Case 3: IconName = "shaman"
                        Case 4: IconName = "Spirit Walker"
                        Case 5: IconName = "Shadow Hunter"
                        Case 6: IconName = "Rexxar"
                    End Select
                End If

            Case "R" ' Random
                ' ...
                If (game = "WAR3") Then
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "dragon whelp"
                        Case 3: IconName = "blue dragon"
                        Case 4: IconName = "red dragon"
                        Case 5: IconName = "Deathwing"
                    End Select
                Else
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "myrmidon"
                        Case 3: IconName = "siren"
                        Case 4: IconName = "dragon turtle"
                        Case 5: IconName = "sea witch"
                        Case 6: IconName = "Illidan"
                    End Select
                End If

            Case "D" ' Tournament
                ' ...
                If (game = "W3XP") Then
                    ' ...
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "Felguard"
                        Case 3: IconName = "infernal"
                        Case 4: IconName = "doomguard"
                        Case 5: IconName = "pit lord"
                        Case 6: IconName = "Archimonde"
                    End Select
                End If
        End Select
    End If

End Property

Public Property Let Statstring(str As String)

    ' ...
    If (Len(str) < 6) Then
        Exit Property
    End If

    ' ...
    game = StrReverse$(Left$(str, 4))

    ' ...
    m_stat_string = Mid$(str, 5)

    ' ...
    Call Parse

End Property

Public Property Get Statstring() As String

    ' ...
    Statstring = m_stat_string

End Property

Public Property Get IsSpawn() As Boolean

    ' ...
    IsSpawn = m_spawn

End Property

Public Property Let IsSpawn(bln As Boolean)

    ' ...
    m_spawn = bln

End Property

Public Property Get Clan() As String

    ' ...
    Clan = m_clan

End Property

Public Property Let Clan(str As String)

    ' ...
    m_clan = str

End Property

Public Property Get Wins() As Long

    ' ...
    Wins = m_wins

End Property

Public Property Let Wins(lng As Long)

    ' ...
    m_wins = lng

End Property

Public Property Get LadderRating() As Long

    ' ...
    LadderRating = m_ladder_rating

End Property

Public Property Let LadderRating(lng As Long)

    ' ...
    m_ladder_rating = lng

End Property

Public Property Get HighLadderRating() As Long

    ' ...
    HighLadderRating = m_high_rating

End Property

Public Property Let HighLadderRating(lng As Long)

    ' ...
    m_high_rating = lng

End Property

Public Property Get LadderRank() As Long

    ' ...
    LadderRank = m_ladder_rank

End Property

Public Property Let LadderRank(lng As Long)

    ' ...
    m_ladder_rank = lng

End Property

Public Property Get Level() As Long

    ' ...
    Level = m_level

End Property

Public Property Let Level(lng As Long)

    ' ...
    m_level = lng

End Property

Public Property Get CharacterName() As String

    ' ...
    CharacterName = m_character_name

End Property

Public Property Let CharacterName(str As String)

    ' ...
    m_character_name = str

End Property

Public Property Get CharacterClassID() As Integer

    ' ...
    CharacterClassID = m_character_class_id

End Property

Public Property Let CharacterClassID(i As Integer)

    ' ...
    m_character_class_id = i

End Property

Public Property Get CharacterClass() As String
    
    ' ...
    If ((game = "D2DV") Or (game = "D2XP")) Then
    
        Dim DIIClasses(0 To 7) As String ' ...
        
        DIIClasses(0) = "Amazon"
        DIIClasses(1) = "Sorceress"
        DIIClasses(2) = "Necromancer"
        DIIClasses(3) = "Paladin"
        DIIClasses(4) = "Barbarian"
        DIIClasses(5) = "Druid"
        DIIClasses(6) = "Assassin"
    
        ' ...
        If (CharacterClassID - 1 <= UBound(DIIClasses)) Then
            CharacterClass = DIIClasses(CharacterClassID - 1)
        End If
        
    End If

End Property

Public Property Let CharacterFlags(lng As Long)

    ' ...
    m_character_flags = lng

End Property

Public Property Get CharacterFlags() As Long

    ' ...
    CharacterFlags = m_character_flags

End Property

Public Property Get IsHardcoreCharacter() As Boolean

    ' ...
    IsHardcoreCharacter = ((CharacterFlags And &H4) = &H4)

End Property

Public Property Get IsFemaleCharacter() As Boolean

    ' ...
    IsFemaleCharacter = _
        ((m_character_class_id = 1) Or (m_character_class_id = 2) Or _
            (m_character_class_id = 7))

End Property

Public Property Get IsExpansionCharacter() As Boolean

    ' ...
    IsExpansionCharacter = ((CharacterFlags And &H20) = &H20)

End Property

Public Property Get IsCharacterDead() As Boolean

    ' ...
    IsCharacterDead = ((IsHardcoreCharacter) And ((CharacterFlags And &H8) = &H8))

End Property

Public Property Get IsLadderCharacter() As Boolean

    ' ...
    IsLadderCharacter = ((CharacterFlags And &H40) = &H40)

End Property

Public Property Get ActsCompleted() As Integer

    ' ...
    ActsCompleted = m_acts_completed

End Property

Public Property Let ActsCompleted(i As Integer)

    ' ...
    m_acts_completed = i

End Property

Public Property Get CurrentAct() As Integer

    ' ...
    If ((game = "D2DV") Or (game = "D2XP")) Then
        ' ...
        If (IsExpansionCharacter = False) Then
            CurrentAct = ((ActsCompleted Mod 4) + 1)
        Else
            CurrentAct = ((ActsCompleted Mod 5) + 1)
        End If
    End If

End Property

Public Property Get CurrentDifficulty() As String

    ' ...
    If ((game = "D2DV") Or (game = "D2XP")) Then
        ' ...
        Select Case (CurrentDifficultyID)
            Case 1: CurrentDifficulty = "Normal"
            Case 2: CurrentDifficulty = "Nightmare"
            Case 3: CurrentDifficulty = "Hell"
        End Select
    End If

End Property

Public Property Get CurrentDifficultyID() As Integer

    Dim difficulty As String ' ...
    
    ' ...
    If ((game = "D2DV") Or (game = "D2XP")) Then
        ' ...
        If (IsExpansionCharacter = False) Then
            difficulty = ((ActsCompleted / 4) + 1)
        Else
            difficulty = ((ActsCompleted / 5) + 1)
        End If
    End If
    
    ' ...
    CurrentDifficultyID = Int(difficulty)
    
End Property

Public Property Get CharacterTitle() As String

    ' thanks c0ol for multi-dimensional array idea
    Dim Classic(2, 3, 2)   As String ' ...
    Dim Expansion(2, 3, 2) As String ' ...
    
    ' softcore
    Classic(0, 0, 0) = "Sir"
    Classic(0, 0, 1) = "Dame"
    Classic(0, 1, 0) = "Lord"
    Classic(0, 1, 1) = "Lady"
    Classic(0, 2, 0) = "Baron"
    Classic(0, 2, 1) = "Baroness"
    
    ' hardcore
    Classic(1, 0, 0) = "Count"
    Classic(1, 0, 1) = "Countess"
    Classic(1, 1, 0) = "Duke"
    Classic(1, 1, 1) = "Duchess"
    Classic(1, 2, 0) = "King"
    Classic(1, 2, 1) = "Queen"
    
    ' softcore
    Expansion(0, 0, 0) = "Slayer"
    Expansion(0, 1, 0) = "Champion"
    Expansion(0, 2, 0) = "Patriarch"
    Expansion(0, 2, 1) = "Matriarch"
    
    ' hardcore
    Expansion(1, 0, 0) = "Destroyer"
    Expansion(1, 1, 0) = "Conquerer"
    Expansion(1, 2, 0) = "Guardian"
    
    ' ...
    If ((game = "D2DV") Or (game = "D2XP")) Then
        ' ...
        If ((CurrentDifficultyID - 1) > 0) Then
            ' ...
            If (game = "D2DV") Then
                ' ...
                CharacterTitle = Classic(IsHardcoreCharacter, (CurrentDifficultyID - 2), IsFemaleCharacter)
                
                ' ...
                If (CharacterTitle = vbNullString) Then
                    CharacterTitle = Classic(IsHardcoreCharacter, (CurrentDifficultyID - 2), False)
                End If
            Else
                ' ...
                CharacterTitle = Expansion(IsHardcoreCharacter, (CurrentDifficultyID - 2), IsFemaleCharacter)
                
                ' ...
                If (CharacterTitle = vbNullString) Then
                    CharacterTitle = Expansion(IsHardcoreCharacter, (CurrentDifficultyID - 2), False)
                End If
            End If
        End If
    End If
    
End Property

Public Property Get Dots() As Long

    ' ...
    Dots = m_dots

End Property

Public Property Let Dots(lng As Long)

    ' ...
    m_dots = lng

End Property

Public Property Get Strength() As Long

    ' ...
    Strength = m_strength

End Property

Public Property Let Strength(lng As Long)

    ' ...
    m_strength = lng

End Property

Public Property Get Magic() As Long

    ' ...
    Magic = m_magic

End Property

Public Property Let Magic(lng As Long)

    ' ...
    m_magic = lng

End Property

Public Property Get Gold() As Long

    ' ...
    Gold = m_gold

End Property

Public Property Let Gold(lng As Long)

    ' ...
    m_gold = lng

End Property

Public Property Get Dexterity() As Long

    ' ...
    Dexterity = m_dexterity

End Property

Public Property Let Dexterity(lng As Long)

    ' ...
    m_dexterity = lng

End Property

Public Property Get Vitality() As Boolean

    ' ...
    Vitality = m_vitality

End Property

Public Property Let Vitality(bln As Boolean)

    ' ...
    m_vitality = bln

End Property

Public Property Get Realm() As String

    ' ...
    Realm = m_realm

End Property

Public Property Let Realm(str As String)

    ' ...
    m_realm = str

End Property

Public Property Get ToString() As String

    ' ...
    If (Statstring = vbNullString) Then
        Exit Sub
    End If
    
    ' ...
    Select Case (game)
        ' ...
        Case "SSHR", "STAR", "JSTR", "SEXP"
            ToString = StarCraft_ToString()
            
        ' ...
        Case "DSHR", "DRTL"
            ToString = Diablo_ToString()
            
        ' ...
        Case "D2DV", "D2XP"
            ToString = DiabloII_ToString()

        ' ...
        Case "W2BN"
            ToString = WarCraftII_ToString()
        
        ' ...
        Case "WAR3", "W3XP"
            ToString = WarCraftIII_ToString()
    End Select

End Property

Private Sub Parse()

    ' ...
    If (Statstring = vbNullString) Then
        Exit Sub
    End If

    ' ...
    Select Case (game)
        ' ...
        Case "SSHR", "STAR", "JSTR", "SEXP"
            Call ParseStarCraft
            
        ' ...
        Case "DSHR", "DRTL"
            Call ParseDiablo
            
        ' ...
        Case "D2DV", "D2XP"
            Call ParseDiabloII

        ' ...
        Case "W2BN"
            Call ParseWarCraftII
        
        ' ...
        Case "WAR3", "W3XP"
            Call ParseWarCraftIII
    End Select

End Sub

Private Sub ParseStarCraft()

    Dim values() As String ' ...

    ' ...
    values() = Split(Statstring, Space$(1))
    
    ' ...
    If (UBound(values) < 7) Then
        Exit Sub
    End If
            
    ' ....
    LadderRating = Val(values(0))
    LadderRank = Val(values(1))
    Wins = Val(values(2))
    IsSpawn = CBool(values(3))
    LadderRank = Val(values(5))
    Icon = values(8)

End Sub

Private Sub ParseDiablo()

    Dim values() As String ' ...

    ' ...
    values() = Split(Statstring, Space$(1))
    
    ' ...
    If (UBound(values) < 7) Then
        Exit Sub
    End If
    
    ' ...
    Level = Val(values(0))
    CharacterClassID = Val(values(1))
    Dots = Val(values(2))
    Strength = Val(values(3))
    Magic = Val(values(4))
    Dexterity = Val(values(5))
    Vitality = Val(values(6))
    Gold = Val(values(7))
    
End Sub

Private Sub ParseDiabloII()

    Dim values()   As String  ' ...
    Dim charData() As Integer ' ...

    ' ...
    values() = Split(Statstring, ",", 3)
    
    ' ...
    If (UBound(values) > 0) Then
        ' ...
        Realm = values(0)
        CharacterName = values(1)
        
        ' ...
        MakeArr values(2), charData
        
        ' ...
        CharacterClassID = charData(13)
        Level = charData(25)
        CharacterFlags = charData(26)
        ActsCompleted = (charData(27) And &H18)
    End If
    
End Sub

Private Sub ParseWarCraftII()

    ' ...
    Call ParseStarCraft

End Sub

Private Sub ParseWarCraftIII()

    Dim values() As String ' ...

    ' ...
    values() = Split(Statstring, Space$(1))
    
    ' ...
    If (UBound(values) < 1) Then
        Exit Sub
    End If
    
    ' ...
    Icon = values(0)
    Level = Val(values(1))
    
    ' ...
    If (UBound(values) > 1) Then
        Clan = StrReverse$(values(2))
    End If

End Sub

Private Function StarCraft_ToString() As String

    ' ...

End Function

Private Function Diablo_ToString() As String

    ' ...

End Function

Private Function DiabloII_ToString() As String

    ' ...

End Function

Private Function WarCraftII_ToString() As String

    ' ...

End Function

Private Function WarCraftIII_ToString() As String

    ' ...

End Function

Private Sub MakeArr(str As String, arr() As Integer)
    
    Dim i As Integer ' ...
    
    ' ...
    For i = 1 To Len(str)
        ' ...
        ReDim Preserve arr(0 To i)
        
        ' ...
        arr(i) = Asc(Mid(str, i, 1))
    Next i
    
End Sub

Private Function MakeLongGameName(game As String) As String

    ' ...
    Select Case (game)
        ' starcraft
        Case "SSHR": MakeLongGameName = "StarCraft: Shareware"
        Case "STAR": MakeLongGameName = "StarCraft"
        Case "JSTR": MakeLongGameName = "StarCraft: Japanese"
        Case "SEXP": MakeLongGameName = "StarCraft: Brood War"

        ' diablo
        Case "DSHR": MakeLongGameName = "Diablo I: Shareware"
        Case "DRTL": MakeLongGameName = "Diablo I: Retail"

        ' diablo ii
        Case "D2DV": MakeLongGameName = "Diablo II"
        Case "D2XP": MakeLongGameName = "Diablo II: Lord of Destruction"

        ' warcraft ii
        Case "W2BN": MakeLongGameName = "WarCraft II"

        ' warcraft iii
        Case "WAR3": MakeLongGameName = "WarCraft III: Reign of Chaos"
        Case "W3XP": MakeLongGameName = "WarCraft III: The Frozen Throne"
    End Select

End Function
