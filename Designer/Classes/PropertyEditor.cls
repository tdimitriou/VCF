VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropertyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Container As PictureBox
Private m_SelectedObject As Object
Private m_Properties As Collection
Private m_ScrollPosition As Long

Public Property Get SelectedObject() As Object
    Set SelectedObject = m_SelectedObject
End Property

Public Property Set SelectedObject(ByVal Value As Object)
    Set m_SelectedObject = Value
    Refresh
End Property

Public Sub Initialize(ByVal Container As PictureBox)
    On Error Resume Next
    
    Set m_Container = Container
    Set m_Properties = New Collection
    
    ' Setup container
    Container.AutoRedraw = True
    Container.BackColor = vbWhite
    Container.ScrollBars = 2 ' Vertical scrollbar
    
    If Err Then Err.Clear
End Sub

Public Sub Refresh()
    On Error Resume Next
    
    If m_Container Is Nothing Then Exit Sub
    
    m_Container.Cls
    
    If m_SelectedObject Is Nothing Then
        DrawEmpty
        Exit Sub
    End If
    
    ' Collect properties
    CollectProperties
    
    ' Draw properties
    DrawProperties
    
    m_Container.Refresh
    
    If Err Then Err.Clear
End Sub

Private Sub DrawEmpty()
    On Error Resume Next
    
    m_Container.CurrentX = 10
    m_Container.CurrentY = 10
    m_Container.Print "No selection"
    
    If Err Then Err.Clear
End Sub

Private Sub CollectProperties()
    On Error Resume Next
    
    Set m_Properties = New Collection
    
    If m_SelectedObject Is Nothing Then Exit Sub
    
    ' Get object type
    Dim ObjType As String
    ObjType = TypeName(m_SelectedObject)
    
    ' Add type name
    m_Properties.Add "Type: " & ObjType, "Type"
    
    ' If it's a UIElement, get common properties
    If TypeOf m_SelectedObject Is VCF.IUIElement Then
        Dim UIElement As VCF.IUIElement
        Set UIElement = m_SelectedObject
        
        AddProperty "Name", UIElement.Name
        AddProperty "DesignLeft", UIElement.DesignLeft
        AddProperty "DesignTop", UIElement.DesignTop
        AddProperty "DesignWidth", UIElement.DesignWidth
        AddProperty "DesignHeight", UIElement.DesignHeight
        AddProperty "Visibility", GetVisibilityString(UIElement.Visibility)
    End If
    
    ' If it's a DependencyObject, get dependency properties
    If TypeOf m_SelectedObject Is VCF.IDependencyObject Then
        Dim DepObj As VCF.IDependencyObject
        Set DepObj = m_SelectedObject
        
        CollectDependencyProperties DepObj
    End If
    
    ' Try to get other common properties
    CollectCommonProperties
    
    If Err Then Err.Clear
End Sub

Private Sub CollectDependencyProperties(ByVal DepObj As IDependencyObject)
    On Error Resume Next
    
    Dim Prop As VCF.DependencyProperty
    Dim PropValue
    Dim PropName As String
    
    For Each Prop In DepObj.DependencyProperties.RegisteredProperties
        PropName = Prop.Name
        
        ' Skip certain properties
        If ShouldSkipProperty(PropName) Then GoTo SKIP_PROP
        
        ' Get value
        VCF.API.CopyVariable Prop.GetValue, PropValue
        
        ' Format value
        AddProperty PropName, FormatPropertyValue(PropValue)
        
SKIP_PROP:
    Next
    
    If Err Then Err.Clear
End Sub

Private Sub CollectCommonProperties()
    On Error Resume Next
    
    Dim PropNames() As String
    Dim i As Long
    
    ' Try to get some common properties via reflection
    PropNames = Split("Text,Content,Title,Caption,BackColor,ForeColor,FontSize,FontBold", ",")
    
    For i = 0 To UBound(PropNames)
        Dim PropValue
        If GetPropertyValue(m_SelectedObject, Trim$(PropNames(i)), PropValue) Then
            AddProperty Trim$(PropNames(i)), FormatPropertyValue(PropValue)
        End If
    Next
    
    If Err Then Err.Clear
End Sub

Private Function GetPropertyValue(ByVal Obj As Object, ByVal PropName As String, ByRef Value As Variant) As Boolean
    On Error Resume Next
    
    Call VCF.API.CopyVariable(CallByName(Obj, PropName, VbGet), Value)
    
    GetPropertyValue = (Err.Number = 0)
    
    If Err Then Err.Clear
End Function

Private Sub AddProperty(ByVal Name As String, ByVal Value As Variant)
    On Error Resume Next
    
    Dim Prop As PropertyItem
    Set Prop = New PropertyItem
    Prop.Name = Name
    Prop.Value = Value
    
    m_Properties.Add Prop, Name
    
    If Err Then Err.Clear
End Sub

Private Function FormatPropertyValue(ByVal Value As Variant) As String
    On Error Resume Next
    
    If IsEmpty(Value) Then
        FormatPropertyValue = "(empty)"
        Exit Function
    End If
    
    If IsNull(Value) Then
        FormatPropertyValue = "(null)"
        Exit Function
    End If
    
    If IsNothing(Value) Then
        FormatPropertyValue = "(nothing)"
        Exit Function
    End If
    
    If IsObject(Value) Then
        FormatPropertyValue = TypeName(Value)
    Else
        FormatPropertyValue = CStr(Value)
    End If
    
    If Err Then FormatPropertyValue = CStr(Value)
End Function

Private Function GetVisibilityString(ByVal Visibility As Long) As String
    On Error Resume Next
    
    Select Case Visibility
        Case 0 ' Visible
            GetVisibilityString = "Visible"
        Case 1 ' Hidden
            GetVisibilityString = "Hidden"
        Case 2 ' Collapsed
            GetVisibilityString = "Collapsed"
        Case Else
            GetVisibilityString = CStr(Visibility)
    End Select
    
    If Err Then GetVisibilityString = CStr(Visibility)
End Function

Private Function ShouldSkipProperty(ByVal PropName As String) As Boolean
    Select Case LCase$(PropName)
        Case "datacontext", "parent", "widget", "widgets", "children"
            ShouldSkipProperty = True
        Case Else
            ShouldSkipProperty = False
    End Select
End Function

Private Sub DrawProperties()
    On Error Resume Next
    
    If m_Properties Is Nothing Then Exit Sub
    
    Dim Y As Long
    Dim LineHeight As Long
    LineHeight = 20
    Y = 10 - m_ScrollPosition
    
    Dim Prop As PropertyItem
    For Each Prop In m_Properties
        ' Draw property name
        m_Container.CurrentX = 5
        m_Container.CurrentY = Y
        m_Container.ForeColor = vbBlue
        m_Container.Print Prop.Name & ":"
        
        ' Draw property value
        Dim ValueX As Long
        ValueX = 100
        m_Container.CurrentX = ValueX
        m_Container.ForeColor = vbBlack
        m_Container.Print Prop.Value
        
        Y = Y + LineHeight
        
        ' Check if we need to scroll
        If Y > m_Container.ScaleHeight Then Exit For
    Next
    
    If Err Then Err.Clear
End Sub

Public Sub SetPropertyValue(ByVal PropName As String, ByVal Value As Variant)
    On Error Resume Next
    
    If m_SelectedObject Is Nothing Then Exit Sub
    
    ' Try to set the property
    If TypeOf m_SelectedObject Is VCF.IDependencyObject Then
        Dim DepObj As VCF.IDependencyObject
        Set DepObj = m_SelectedObject
        
        If DepObj.DependencyProperties.Exists(PropName) Then
            DepObj.DependencyProperties.SetValue PropName, Value
        Else
            ' Try regular property
            If IsObject(Value) Then
                CallByName m_SelectedObject, PropName, VbSet, Value
            Else
                CallByName m_SelectedObject, PropName, VbLet, Value
            End If
        End If
    Else
        ' Try regular property
        If IsObject(Value) Then
            CallByName m_SelectedObject, PropName, VbSet, Value
        Else
            CallByName m_SelectedObject, PropName, VbLet, Value
        End If
    End If
    
    ' Refresh display
    Refresh
    
    If Err Then
        MsgBox "Error setting property: " & Err.Description, vbExclamation
    End If
End Sub

Public Sub Scroll(ByVal Delta As Long)
    On Error Resume Next
    
    m_ScrollPosition = m_ScrollPosition + Delta
    If m_ScrollPosition < 0 Then m_ScrollPosition = 0
    
    Refresh
    
    If Err Then Err.Clear
End Sub

