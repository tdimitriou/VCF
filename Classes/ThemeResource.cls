VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThemeResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMarkupExtension

Public Key As String
Attribute Key.VB_VarUserMemId = 0

Private Function IMarkupExtension_ProvideValue() As Variant
    VCF.API.CopyVariable ProvideValue, IMarkupExtension_ProvideValue
End Function

Public Function ProvideValue() As Variant
    Dim Theme As Object
    
    VCF.API.CopyVariable GetThemeResource, ProvideValue
End Function

Private Function GetThemeResource()
Try:
    On Error GoTo Catch
    
    Dim ActiveTheme As ObservableDictionary
    Set ActiveTheme = ThemesManager.ActiveTheme
    
    If ActiveTheme Is Nothing Then Exit Function
    
    API.CopyVariable ActiveTheme(Key), GetThemeResource
    
    Exit Function
    
Catch:
    
    GetThemeResource = Empty
    
End Function

Private Property Get ThemesManager() As ThemesManager
Try:
    On Error GoTo Catch
        
    Set ThemesManager = Application.Current.Base.ThemesManager
            
    Exit Property
    
Catch:
    
    Set ThemesManager = Nothing
    
End Property

