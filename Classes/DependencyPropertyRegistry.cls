VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DependencyPropertyRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_TypeDefs As cSortedDictionary
Private m_Attached As cSortedDictionary
Private m_AttachedOwners As cSortedDictionary
Private m_Initialized As Boolean

Private Sub Class_Initialize()
    Set m_TypeDefs = New_c.SortedDictionary(TextCompare)
    Set m_Attached = New_c.SortedDictionary(TextCompare)
    Set m_AttachedOwners = New_c.SortedDictionary(TextCompare)
End Sub

Public Sub EnsureBuiltInTypes()
    If m_Initialized Then Exit Sub
    RegisterFrameworkElementType
    RegisterBuiltInAttached
    m_Initialized = True
End Sub

' WPF-shaped attached metadata (storage remains AttachedProperties dict).
Public Sub RegisterAttached( _
    ByVal OwnerType As String, _
    ByVal Name As String, _
    ByVal PropertyType As VbVarType, _
    Optional ByVal DefaultValue As Variant)

    Dim FullName As String

    If Len(OwnerType) = 0 Or Len(Name) = 0 Then Exit Sub
    FullName = OwnerType & "." & Name

    If m_Attached.Exists(FullName) Then m_Attached.Remove FullName
    m_Attached.Add FullName, Array(OwnerType, Name, PropertyType, DefaultValue)

    If Not m_AttachedOwners.Exists(OwnerType) Then
        m_AttachedOwners.Add OwnerType, True
    End If
End Sub

Public Function IsAttachedRegistered(ByVal FullName As String) As Boolean
    EnsureBuiltInTypes
    IsAttachedRegistered = m_Attached.Exists(FullName)
End Function

Public Function HasAttachedOwner(ByVal OwnerType As String) As Boolean
    EnsureBuiltInTypes
    HasAttachedOwner = m_AttachedOwners.Exists(OwnerType)
End Function

Public Function GetAttachedDefault(ByVal FullName As String) As Variant
    Dim Entry As Variant

    EnsureBuiltInTypes
    If Not m_Attached.Exists(FullName) Then Exit Function
    Entry = m_Attached(FullName)
    GetAttachedDefault = Entry(3)
End Function

Public Sub RegisterType( _
    ByVal TypeName As String, _
    ByVal BaseTypeName As String, _
    ParamArray PropertyDefs())

    Dim Defs As cCollection
    Dim i As Long
    Dim Entry As Variant

    If m_TypeDefs.Exists(TypeName) Then
        Set Defs = m_TypeDefs(TypeName)
    Else
        Set Defs = New_c.Collection
        m_TypeDefs.Add TypeName, Defs
    End If

    If Len(BaseTypeName) > 0 Then
        ApplyInheritedDefs Defs, BaseTypeName
    End If

    For i = LBound(PropertyDefs) To UBound(PropertyDefs)
        Entry = PropertyDefs(i)
        AddOrReplaceDef Defs, Entry
    Next
End Sub

Public Sub ApplyTo(ByVal TypeName As String, ByVal Props As DependencyProperties)
    Dim Defs As cCollection
    Dim Entry As Variant
    Dim Meta As DependencyPropertyMetadata
    Dim PropName As String
    Dim PropType As VbVarType
    Dim PropTypeName As String
    Dim ProgId As String
    Dim BindingModeValue As BindingMode

    EnsureBuiltInTypes

    If m_TypeDefs.Exists(TypeName) Then
        Set Defs = m_TypeDefs(TypeName)
    ElseIf m_TypeDefs.Exists("FrameworkElement") Then
        Set Defs = m_TypeDefs("FrameworkElement")
    Else
        Exit Sub
    End If

    For Each Entry In Defs
        PropName = CStr(Entry(0))
        PropTypeName = vbNullString
        ProgId = vbNullString
        If Not Props.Exists(PropName) Then
            PropType = Entry(1)
            If Len(CStr(Entry(2))) > 0 Then PropTypeName = CStr(Entry(2))
            If Len(CStr(Entry(3))) > 0 Then ProgId = CStr(Entry(3))
            BindingModeValue = Entry(7)
            Set Meta = NewDependencyPropertyMetadata(CBool(Entry(4)), CBool(Entry(5)), CBool(Entry(6)), BindingModeValue)
            Call Props.Register(PropName, PropType, PropTypeName, ProgId, , Meta)
        End If
    Next
End Sub

Public Function HasType(ByVal TypeName As String) As Boolean
    HasType = m_TypeDefs.Exists(TypeName)
End Function

Private Sub RegisterFrameworkElementType()
    RegisterType "FrameworkElement", vbNullString, _
        Array("Width", vbDouble, vbNullString, vbNullString, True, False, False, OneWay), _
        Array("Height", vbDouble, vbNullString, vbNullString, True, False, False, OneWay), _
        Array("Margin", vbObject, "Thickness", vbNullString, True, False, False, OneWay), _
        Array("Visibility", vbLong, vbNullString, vbNullString, False, True, False, OneWay), _
        Array("MinWidth", vbDouble, vbNullString, vbNullString, True, False, False, OneWay), _
        Array("MinHeight", vbDouble, vbNullString, vbNullString, True, False, False, OneWay), _
        Array("MaxWidth", vbDouble, vbNullString, vbNullString, True, False, False, OneWay), _
        Array("MaxHeight", vbDouble, vbNullString, vbNullString, True, False, False, OneWay)

    RegisterType "Panel", "FrameworkElement", _
        Array("ShowGridLines", vbBoolean, vbNullString, vbNullString, False, True, False, OneWay)

    RegisterType "StackPanel", "FrameworkElement", _
        Array("Orientation", vbString, vbNullString, vbNullString, False, True, False, OneWay)
    RegisterType "Grid", "FrameworkElement"
    RegisterType "ContentControl", "FrameworkElement", _
        Array("Content", vbVariant, vbNullString, vbNullString, False, True, False, OneWay), _
        Array("ContentTemplate", vbObject, "DataTemplate", vbNullString, False, True, False, OneWay)

    RegisterType "ItemsControl", "FrameworkElement", _
        Array("ItemsSource", vbObject, vbNullString, vbNullString, False, True, False, OneWay), _
        Array("ItemTemplate", vbObject, "DataTemplate", vbNullString, False, True, False, OneWay)

    RegisterType "Selector", "ItemsControl", _
        Array("SelectedItem", vbObject, vbNullString, vbNullString, False, True, False, TwoWay), _
        Array("SelectedIndex", vbLong, vbNullString, vbNullString, False, True, False, TwoWay), _
        Array("SelectedValue", vbVariant, vbNullString, vbNullString, False, True, False, TwoWay), _
        Array("SelectedValuePath", vbString, vbNullString, vbNullString, False, True, False, OneWay)

    RegisterType "ListView", "Selector", _
        Array("Margin", vbObject, "Thickness", vbNullString, True, False, False, OneWay), _
        Array("Padding", vbObject, "Thickness", vbNullString, True, False, False, OneWay)

    RegisterType "Border", "FrameworkElement", _
        Array("Child", vbObject, vbNullString, vbNullString, False, True, False, OneWay)
    RegisterType "UserControl", "FrameworkElement"
    RegisterType "Button", "FrameworkElement", _
        Array("Content", vbVariant, vbNullString, vbNullString, False, True, False, OneWay), _
        Array("Padding", vbObject, "Thickness", vbNullString, True, False, False, OneWay)
    RegisterType "Window", "FrameworkElement"
End Sub

Private Sub RegisterBuiltInAttached()
    RegisterAttached "Grid", "Row", vbLong, 0&
    RegisterAttached "Grid", "Column", vbLong, 0&
    RegisterAttached "Grid", "RowSpan", vbLong, 1&
    RegisterAttached "Grid", "ColumnSpan", vbLong, 1&
End Sub

Private Sub ApplyInheritedDefs(ByVal Defs As cCollection, ByVal BaseTypeName As String)
    Dim BaseDefs As cCollection
    Dim Entry As Variant
    Dim Name As String
    Dim Exists As Boolean

    If Not m_TypeDefs.Exists(BaseTypeName) Then Exit Sub
    Set BaseDefs = m_TypeDefs(BaseTypeName)

    For Each Entry In BaseDefs
        Name = Entry(0)
        Exists = False
        Dim Existing As Variant
        For Each Existing In Defs
            If Existing(0) = Name Then Exists = True: Exit For
        Next
        If Not Exists Then Defs.Add Entry
    Next
End Sub

Private Sub AddOrReplaceDef(ByVal Defs As cCollection, ByVal Entry As Variant)
    Dim i As Long
    For i = 1 To Defs.Count
        If Defs(i)(0) = Entry(0) Then
            Defs.Remove i
            Exit For
        End If
    Next
    Defs.Add Entry
End Sub
