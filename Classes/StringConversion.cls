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
    End Function
    
    Public Function ToLowerCase(ByVal Value As String)
        ToLowerCase = StrConv(Value, vbLowerCase, m_LocaleID)
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
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        m_LocaleID = 1033
                
        Set AccentedCharacters = New_c.SortedDictionary
        With AccentedCharacters
            .Add "¢", "Á"       ' 'A
            .Add "¸", "Å"       ' 'E
            .Add "¹", "Ç"       ' 'H
            .Add "º", "É"       ' 'I
            .Add "¼", "Ï"       ' 'O
            .Add "¾", "Õ"       ' 'Y
            .Add "¿", "Ù"       ' Capital 'Omega' with accent
            .Add "ò", "Ó"       ' Final 'Sigma' (put in accented characters as an exception
                                ' because it must always been replaced in Uppercase,
                                ' (the same rule as accented chars in Uppercase)
        End With
    End Sub
'#End Region


