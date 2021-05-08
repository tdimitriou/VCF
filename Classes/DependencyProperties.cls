VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DependencyProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Props As cSortedDictionary
Private m_TargetPtr As Long

Public Event DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue)

'#Region Public Members
    Public Sub Register(PropertyName As String, _
                        PropertyType As VbVarType, _
                        Optional PropertyTypeName As String, _
                        Optional ProgId As String, _
                        Optional UnsetValue, _
                        Optional Metadata As DependencyPropertyMetadata)
        
        Dim p As DependencyProperty
        Set p = New DependencyProperty
        
                       
        Call p.Register(ObjPtr(Me), PropertyName, PropertyType, PropertyTypeName, ProgId, UnsetValue, Metadata)
        
        m_Props.Add p.Name, p
    End Sub

    Public Sub SetCurrentValue(ByVal PropertyName As String, Value)
        Dim p As DependencyProperty
        Set p = m_Props(PropertyName)
        
        p.SetCurrentValue Value
    End Sub
    
    Public Sub SetValue(ByVal PropertyName As String, Value)
        Dim p As DependencyProperty
        Set p = m_Props(PropertyName)
        
        p.SetValue Value
    End Sub
        
    Public Function GetValue(ByVal PropertyName As String)
        Dim p As DependencyProperty
        
        Set p = m_Props(PropertyName)
        
        Call API.CopyVariable(p.GetValue, GetValue)
        
'        If Equals(GetValue, p.UnsetValue) Then
'            ' Return the value of the parent object
'            If Target Is Nothing Then Exit Function
'            If Target.Parent Is Nothing Then Exit Function
'
'            If Target.Parent.DependencyProperties.Exists(PropertyName) Then
'                Call API.CopyVariable(Target.Parent.DependencyProperties.GetValue(PropertyName), GetValue)
'            End If
'        End If
    End Function
    
    Public Function Exists(PropertyName As String) As Boolean
        Exists = m_Props.Exists(PropertyName)
    End Function
    
    Public Function GetProperty(PropertyName As String) As DependencyProperty
        Set GetProperty = m_Props.Item(PropertyName)
    End Function
'#End Region

'#Region Friend Members
    Friend Property Get Target() As IDependencyObject
        Dim o As Object
        Set o = API.ObjFromPtr(m_TargetPtr)
        If o Is Nothing Then Exit Property
        If Not TypeOf o Is IDependencyObject Then Exit Property
        
        Set Target = o
    End Property
    
    Friend Sub OnDependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue)
        RaiseEvent DependencyPropertyChanged(Property, PreviousValue)
    End Sub
    
    Friend Property Get RegisteredProperties() As cSortedDictionary
        Set RegisteredProperties = m_Props
    End Property
'#End Region

'#Region Constructors
    Friend Function DependencyProperties(Target As Object) As DependencyProperties
        If Not TypeOf Target Is IDependencyObject Then Exit Function
        m_TargetPtr = ObjPtr(Target)
        
        Set DependencyProperties = Me
    End Function
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Props = New_c.SortedDictionary(TextCompare)
    End Sub
'#End Region
