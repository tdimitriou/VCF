VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HarnessScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VCF.IDependencyObject
Implements VCF.IControl
Implements VCF.IUIElement
Implements VCF.IUserControl

Private m_Base As VCF.UserControl
Private m_BaseUIElement As UIElementBase
Private m_XamlKey As String
Private m_WireLoginOnActivate As Boolean
Private m_ViewLoaded As Boolean

Friend Sub Configure(ByVal XamlResourceKey As String, ByVal WireLoginOnActivate As Boolean)
    m_XamlKey = XamlResourceKey
    m_WireLoginOnActivate = WireLoginOnActivate
End Sub

Public Property Get XamlResourceKey() As String
    XamlResourceKey = m_XamlKey
End Property

Public Property Get IsViewLoaded() As Boolean
    IsViewLoaded = m_ViewLoaded
End Property

Public Sub LoadView()
    If m_ViewLoaded Then Exit Sub
    If Len(m_XamlKey) = 0 Then Err.Raise vbObjectError, , "HarnessScreen not configured"
    InitializeComponent
    m_ViewLoaded = True
End Sub

'#Region IDependencyObject Inheritance
    Public Property Get DependencyProperties() As DependencyProperties
        Set DependencyProperties = m_Base.DependencyProperties
    End Property
'#End Region

'#Region IDependencyObject Interface Implementation
    Private Property Get IDependencyObject_Children() As Object
        Set IDependencyObject_Children = Children
    End Property
    
    Private Property Get IDependencyObject_DependencyProperties() As DependencyProperties
        Set IDependencyObject_DependencyProperties = DependencyProperties
    End Property

    Private Property Get IDependencyObject_Parent() As IDependencyObject
        Set IDependencyObject_Parent = Parent
    End Property
'#End Region

'#Region IUIElement Inheritance
    Public Property Get AttachedProperties() As VCF.ObservableDictionary
        Set AttachedProperties = m_BaseUIElement.AttachedProperties
    End Property
    
    Public Property Set DataContext(ByVal Value As Object)
        Set m_Base.DataContext = Value
        
        ' TO-DO: Recreate the Bindings!!!
    End Property
    
    Public Property Get DataContext() As Object
        Set DataContext = m_Base.DataContext
    End Property
    
    Public Property Let DesignHeight(ByVal Value As Double)
        m_Base.DesignHeight = Value
    End Property
    
    Public Property Get DesignHeight() As Double
        DesignHeight = m_Base.DesignHeight
    End Property
    
    Public Property Let DesignLeft(ByVal Value As Double)
        m_Base.DesignLeft = Value
    End Property
    
    Public Property Get DesignLeft() As Double
        DesignLeft = m_Base.DesignLeft
    End Property
    
    Public Property Let DesignTop(ByVal Value As Double)
        m_Base.DesignTop = Value
    End Property
    
    Public Property Get DesignTop() As Double
        DesignTop = m_Base.DesignTop
    End Property
    
    Public Property Let DesignWidth(ByVal Value As Double)
        m_Base.DesignWidth = Value
    End Property
    
    Public Property Get DesignWidth() As Double
        DesignWidth = m_Base.DesignWidth
    End Property

    Public Property Let Visibility(ByVal Value As Visibility)
        m_Base.Visibility = Value
    End Property
    
    Public Property Get Visibility() As Visibility
        Visibility = m_Base.Visibility
    End Property

    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        m_Base.Move Left, Top, Width, Height
    End Sub

    Public Property Set Parent(ByVal Value As IControl)
        Set m_Base.Parent = Value
    End Property
    
    Public Property Get Parent() As IControl
        Set Parent = m_Base.Parent
    End Property
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        Call VCF.NewUserControl(Me, m_Base)
        NewUIElementBase Me, m_BaseUIElement
    End Sub
'#End Region

'#Region IControl Interface Implementation
    Private Property Get IControl_Children() As UIElementCollection
        Set IControl_Children = Children
    End Property
    
    Private Property Get IControl_Widget() As vbRichClient5.cWidgetBase
        Set IControl_Widget = Widget
    End Property
    
    Private Property Get IControl_Widgets() As vbRichClient5.cWidgets
        Set IControl_Widgets = Widgets
    End Property
'#End Region

'#Region IControl Inheritance
    Public Property Get Children() As UIElementCollection
        Set Children = m_Base.Children
    End Property
    
    Public Property Get Widget() As cWidgetBase
        Set Widget = m_Base.Widget
    End Property
    
    Public Property Get Widgets() As cWidgets
        Set Widgets = m_Base.Widgets
    End Property
'#End Region

'#Region IUIElement Interface Implementation
    Private Property Get IUIElement_AttachedProperties() As VCF.ObservableDictionary
        Set IUIElement_AttachedProperties = m_BaseUIElement.AttachedProperties
    End Property
    
    Private Property Get IUIElement_Base() As VCF.UIElementBase
        Set IUIElement_Base = m_BaseUIElement
    End Property

    Private Property Set IUIElement_DataContext(ByVal RHS As Object)
        Set DataContext = RHS
    End Property
    
    Private Property Get IUIElement_DataContext() As Object
        Set IUIElement_DataContext = DataContext
    End Property
        
    Private Property Let IUIElement_DesignHeight(ByVal RHS As Double)
        DesignHeight = RHS
    End Property
    
    Private Property Get IUIElement_DesignHeight() As Double
        IUIElement_DesignHeight = DesignHeight
    End Property
    
    Private Property Let IUIElement_DesignLeft(ByVal RHS As Double)
        DesignLeft = RHS
    End Property
    
    Private Property Get IUIElement_DesignLeft() As Double
        IUIElement_DesignLeft = DesignLeft
    End Property
    
    Private Property Let IUIElement_DesignTop(ByVal RHS As Double)
        DesignTop = RHS
    End Property
    
    Private Property Get IUIElement_DesignTop() As Double
        IUIElement_DesignTop = DesignTop
    End Property
    
    Private Property Let IUIElement_DesignWidth(ByVal RHS As Double)
        DesignWidth = RHS
    End Property
    
    Private Property Get IUIElement_DesignWidth() As Double
        IUIElement_DesignWidth = DesignWidth
    End Property
    
    Private Sub IUIElement_Move(Left As Single, Optional Top As Single = 0#, Optional Width As Single, Optional Height As Single)
        Move Left, Top, Width, Height
    End Sub

    Private Property Set IUIElement_Parent(RHS As IControl)
        Set Parent = RHS
    End Property
    
    Private Property Get IUIElement_Parent() As IControl
        Set IUIElement_Parent = Parent
    End Property
'#End Region

'#Region IUserControl Inheritance
    Public Property Get Base() As VCF.UserControl
        Set Base = m_Base
    End Property
    
    Public Property Get Name() As String
        Name = m_Base.Name
    End Property
    
    Public Property Let Name(ByVal Value As String)
        m_Base.Name = Value
    End Property
'#End Region

'#Region IUserControl Interface Implementation
    Private Property Get IUserControl_Base() As VCF.UserControl
        Set IUserControl_Base = Me.Base
    End Property
    
    Private Property Get IUserControl_DependencyProperties() As DependencyProperties
        Set IUserControl_DependencyProperties = DependencyProperties
    End Property
    
    Private Property Set IUserControl_DataContext(ByVal Value As Object)
        Set DataContext = Value
    End Property
    
    Private Property Get IUserControl_DataContext() As Object
        Set IUserControl_DataContext = DataContext
    End Property
    
    Private Property Let IUserControl_DesignHeight(ByVal Value As Double)
        DesignHeight = Value
    End Property
    
    Private Property Get IUserControl_DesignHeight() As Double
        IUserControl_DesignHeight = DesignHeight
    End Property
    
    Private Property Let IUserControl_DesignLeft(ByVal Value As Double)
        DesignLeft = Value
    End Property
    
    Private Property Get IUserControl_DesignLeft() As Double
        IUserControl_DesignLeft = DesignLeft
    End Property
    
    Private Property Let IUserControl_DesignTop(ByVal Value As Double)
        DesignTop = Value
    End Property
    
    Private Property Get IUserControl_DesignTop() As Double
        IUserControl_DesignTop = DesignTop
    End Property
    
    Private Property Let IUserControl_DesignWidth(ByVal Value As Double)
        DesignWidth = Value
    End Property
    
    Private Property Get IUserControl_DesignWidth() As Double
        IUserControl_DesignWidth = DesignWidth
    End Property
    
    Private Sub IUserControl_InitializeComponent()
        ' UserControl.Initialize runs during Class_Initialize — before Configure/LoadView.
        If Len(m_XamlKey) = 0 Then Exit Sub
        If m_ViewLoaded Then Exit Sub
        InitializeComponent
        m_ViewLoaded = True
    End Sub

    Private Property Let IUserControl_Name(ByVal Value As String)
        Name = Value
    End Property
    
    Private Property Get IUserControl_Name() As String
        IUserControl_Name = Name
    End Property

    Private Property Let IUserControl_Visibility(ByVal Value As Visibility)
        Visibility = Value
    End Property
    
    Private Property Get IUserControl_Visibility() As Visibility
        IUserControl_Visibility = Visibility
    End Property
    
    Private Sub IUserControl_Move(ByRef Left As Single, ByRef Top As Single, ByRef Width As Single, ByRef Height As Single)
        Move Left, Top, Width, Height
    End Sub
    
    Private Property Set IUserControl_Parent(ByVal Value As IControl)
        Set Parent = Value
    End Property
    
    Private Property Get IUserControl_Parent() As IControl
        Set IUserControl_Parent = Parent
    End Property
    
    Private Property Get IUserControl_Children() As UIElementCollection
        Set IUserControl_Children = Children
    End Property
    
    Private Property Get IUserControl_Widget() As cWidgetBase
        Set IUserControl_Widget = Widget
    End Property
    
    Private Property Get IUserControl_Widgets() As cWidgets
        Set IUserControl_Widgets = Widgets
    End Property
'#End Region

'#Region Private Members
    Public Sub OnActivated()
        On Error Resume Next
        If m_WireLoginOnActivate Then
            WireLoginActivate
        ElseIf InStr(1, m_XamlKey, "MainMenu", vbTextCompare) > 0 Then
            WireMainMenuActivate
        ElseIf InStr(1, m_XamlKey, "SalesOrder", vbTextCompare) > 0 Then
            WireSalesOrderActivate
        End If
        
        If m_Base.Widget.Width > 0 And m_Base.Widget.Height > 0 Then
            Move 0, 0, m_Base.Widget.Width, m_Base.Widget.Height
        End If
        m_Base.Widget.Refresh
    End Sub

    Private Sub WireLoginActivate()
        With VCF.NamingManager.GetNamedChildren(m_Base)
            If .ContainsKey("MyList") Then
                Dim LV As VCF.ListView
                Dim VM As StubLoginViewModel
                Set LV = .Item("MyList")
                Set VM = DataContext
                If Not VM Is Nothing Then
                    Set modApp.ActiveLoginViewModel = VM
                    Set LV.ItemsSource = VM.UserList
                    If VM.UserList.Count > 0 Then LV.SelectedIndex = 0
                    LV.Refresh
                End If
            End If
            If .ContainsKey("Pwd") Then
                Dim Pwd As VCF.TextBox
                Set Pwd = .Item("Pwd")
                Pwd.Widget.SetFocus
            End If
        End With
    End Sub

    Private Sub WireMainMenuActivate()
        ' Milestone 2 — named lookup after RebuildNamedItemsList (HarnessNavigation).
        With VCF.NamingManager.GetNamedChildren(m_Base)
            If .ContainsKey("MenuTitle") Then
                Dim TitleTb As VCF.TextBlock
                Set TitleTb = .Item("MenuTitle")
                Debug.Print "[P7d-MAINMENU] named MenuTitle.Text=" & TitleTb.Text
            End If
            If .ContainsKey("WelcomeText") Then
                Dim WelcomeTb As VCF.TextBlock
                Set WelcomeTb = .Item("WelcomeText")
                Debug.Print "[P7d-MAINMENU] named WelcomeText.Text=" & WelcomeTb.Text
            End If
        End With
    End Sub

    Private Sub WireSalesOrderActivate()
        Dim Vm As StubSalesOrderViewModel
        Dim LV As VCF.ListView
        
        On Error Resume Next
        Set Vm = DataContext
        If Not Vm Is Nothing Then Vm.EnsureSeedLines
        
        With VCF.NamingManager.GetNamedChildren(m_Base)
            If .ContainsKey("LinesList") Then
                Set LV = .Item("LinesList")
                If Not LV Is Nothing And Not Vm Is Nothing Then
                    Set LV.ItemsSource = Vm.Lines
                    LV.Refresh
                End If
            End If
            If .ContainsKey("SalesTitle") Then
                Dim TitleTb As VCF.TextBlock
                Set TitleTb = .Item("SalesTitle")
                Debug.Print "[P7d-SALES] named SalesTitle.Text=" & TitleTb.Text
            End If
            If .ContainsKey("LeftColumn") Then
                Debug.Print "[P7d-SALES] named LeftColumn ok"
            End If
            If .ContainsKey("RightColumn") Then
                Debug.Print "[P7d-SALES] named RightColumn ok"
            End If
            If .ContainsKey("LinesList") And Not Vm Is Nothing Then
                Debug.Print "[P7d-SALES] Lines.Count=" & CStr(Vm.Lines.Count)
            End If
        End With
    End Sub

    Private Sub InitializeComponent()
        Dim Xml As String
        
        If Len(m_XamlKey) = 0 Then Err.Raise vbObjectError, , "HarnessScreen not configured"
        If Not modApp.XamlResources.ContainsKey(m_XamlKey) Then
            Err.Raise vbObjectError, , "Missing XAML resource: " & m_XamlKey
        End If
        
        Xml = AdaptXamlSuperclass(modApp.XamlResources(m_XamlKey))
        
        With New XAMLReader
            Call .LoadSuperclassData(Me, Xml)
        End With
    End Sub
    
    ' LoadSuperclassData requires x:Class = TypeName(Me) ("HarnessScreen"); POS XAML uses SplashView/LoginView.
    Private Function AdaptXamlSuperclass(ByVal Xml As String) As String
        Dim Result As String
        Result = Xml
        Result = Replace(Result, "x:Class=""SplashView""", "x:Class=""HarnessScreen""", , , vbTextCompare)
        Result = Replace(Result, "x:Class=""LoginView""", "x:Class=""HarnessScreen""", , , vbTextCompare)
        Result = Replace(Result, "x:Class=""MainMenuView""", "x:Class=""HarnessScreen""", , , vbTextCompare)
        Result = Replace(Result, "x:Class=""SalesOrderView""", "x:Class=""HarnessScreen""", , , vbTextCompare)
        AdaptXamlSuperclass = Result
    End Function
'#End Region


