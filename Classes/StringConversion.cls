VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_LocaleID As Long

Private PunctuationMarks As cSortedDictionary
Private AccentedCharacters As cSortedDictionary

'#Region Constructors
    Friend Sub Initialize(Optional ByVal LocaleID As Long = 1033)
        m_LocaleID = LocaleID
    End Sub
'#End Region

'#Region Public Properties
    Public Property Get LocaleID() As Long
        LocaleID = m_LocaleID
    End Property
    
    Public Property Let LocaleID(ByVal Value As Long)
        m_LocaleID = Value
    End Property
'#End Region

'#Region Public Methods
    Public Function ToProperCase(ByVal Value As String)
        ToProperCase = StrConv(Value, vbProperCase, m_LocaleID)
        
        If m_LocaleID = 1032 Then ToProperCase = ReplaceGreekFinalSigma(ToProperCase)
    End Function
    
    Public Function ToLowerCase(ByVal Value As String)
        ToLowerCase = StrConv(Value, vbLowerCase, m_LocaleID)
        
        If m_LocaleID = 1032 Then ToLowerCase = ReplaceGreekFinalSigma(ToLowerCase)
    End Function
    
    Public Function ToUpperCase(ByVal Value As String)
        
        ToUpperCase = StrConv(Value, vbUpperCase, m_LocaleID)
        
        If m_LocaleID <> 1032 Then Exit Function
        
        ToUpperCase = RemoveAccentedCharacters(ToUpperCase)
    End Function
'#End Region

'#Region Private Methods
    Private Function RemoveAccentedCharacters(ByVal Value As String) As String
        Dim Index As Long
        
        For Index = 0 To AccentedCharacters.Count - 1
            Value = Replace$(Value, AccentedCharacters.KeyByIndex(Index), AccentedCharacters.ItemByIndex(Index))
        Next
        
        RemoveAccentedCharacters = Value
    End Function
    
    Private Function ReplaceGreekFinalSigma(ByVal Value As String, Optional StartPos As Long = 1) As String
        Dim Pos As Long
    
        Pos = InStr(StartPos, Value, "s", vbBinaryCompare)
        
        If Pos <= 0 Then
            ReplaceGreekFinalSigma = Value
            Exit Function
        End If
        
        If Pos = Len(Value) Then
            Value = Left$(Value, Len(Value) - 1) & "?"
            ReplaceGreekFinalSigma = Value
            Exit Function
        End If
        
        Dim NextChar As String
        NextChar = Mid$(Value, Pos + 1, 1)
        If PunctuationMarks.Exists(NextChar) Then
            Value = Left$(Value, Pos - 1) & "?" & Mid$(Value, Pos + 1, Len(Value) - Pos)
        End If
        
        ReplaceGreekFinalSigma = ReplaceGreekFinalSigma(Value, Pos + 1)
    End Function
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        m_LocaleID = 1033
        
        Set PunctuationMarks = New_c.SortedDictionary(BinaryCompare)
        
        With PunctuationMarks
            .Add " ", " "
            .Add ",", ","
            .Add ".", "."
            .Add "!", "!"
            .Add ";", ";"
            .Add ":", ":"
            .Add "(", "("
            .Add ")", ")"
            .Add "·", "·"
            .Add "-", "-"
            .Add "_", "_"
            .Add "'", "'"
            .Add ""","""
            .Add "«", "«"
            .Add "»", "»"
        End With
        
        Set AccentedCharacters = New_c.SortedDictionary
        With AccentedCharacters
            .Add "?", "?"       ' 'A
            .Add "?", "?"       ' 'E
            .Add "?", "?"       ' 'H
            .Add "?", "?"       ' 'I
            .Add "?", "?"       ' 'O
            .Add "?", "?"       ' 'Y
            .Add "?", "O"       ' Capital 'Omega' with accent
            .Add "?", "S"       ' Final 'Sigma' (put in accented characters as an exception
                                ' because it must always been replaced in Uppercase,
                                ' (the same rule as accented chars in Uppercase)
        End With
    End Sub
'#End Region


