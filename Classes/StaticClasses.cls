VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StaticClasses"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get CollectionViewSource() As CollectionViewSource
    Set CollectionViewSource = modStaticClasses.CollectionViewSource
End Property

Public Property Get Application() As ApplicationStatic
    Set Application = modStaticClasses.Application
End Property

Public Property Get API() As VCF.API
    Set API = modStaticClasses.API
End Property

Public Property Get Object() As ObjectStatic
    Set Object = modStaticClasses.Object
End Property

Public Property Get StringConversion() As StringConversion
    Set StringConversion = modStaticClasses.StringConversion
End Property

Public Function NamingManager() As NamingManager
    Set NamingManager = modStaticClasses.NamingManager
End Function

Public Function StringProcessor() As StringProcessor
    Set StringProcessor = modStaticClasses.StringProcessor
End Function

Public Function Color() As VCF.Color
    Set Color = modStaticClasses.Color
End Function
