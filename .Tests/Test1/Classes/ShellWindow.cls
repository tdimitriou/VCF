VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow

Private Win As Window
Attribute Win.VB_VarHelpID = -1
Private WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1
Private WithEvents NamedItems As ObservableDictionary
Attribute NamedItems.VB_VarHelpID = -1

Private RightColumn As Panel
Private LeftColumn As Panel
Private MyHost As WindowsFormsHost

Public Sub Show()

    'Set Win.DataContext = New SampleDataContext
    
    'Win.Form.Show
    Win.Show
End Sub

Private Sub Class_Initialize()
    NewWindow Me, Win
End Sub

Private Function GetParentWindow(ByVal Obj As IUIElement) As Window
    On Error Resume Next
            
    If Obj Is Nothing Then Exit Function
    
    If TypeOf Obj Is Window Then
        Set GetParentWindow = Obj
    Else
        Dim ObjParent As IUIElement
        Set ObjParent = Obj.Parent
        
        If Not ObjParent Is Nothing Then
            Set GetParentWindow = GetParentWindow(ObjParent)
        End If
    End If
End Function

'CSEH: ErrMsgBox
Public Sub SwapContent()
        '<EhHeader>
        On Error GoTo SwapContent_Err
        '</EhHeader>
        Static bSet As Boolean

100     If (Not bSet) And (Not RightColumn Is Nothing) Then
102         bSet = True
        
104         RightColumn.Widget.LockRefresh = True
106         RightColumn.Children.Clear

108         With New XAMLReader
110             RightColumn.Children.Add .Load(New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\RightColumn.xml"))
            End With

112         RightColumn.Widget.LockRefresh = False
            
            Exit Sub
        End If
    
114     bSet = False
    
        'Debug.Print TypeName(Win.Form.ActiveWidget)
    
    '    Static bExecuted As Boolean
    '    If bExecuted Then Exit Sub
    '    bExecuted = True
        
        Static ViewID As Long
116     ViewID = ViewID + 1
    
    
        Dim ViewPath As String
        Dim Modulo As Long
118     Modulo = ViewID Mod 2
''''        Static Views As cSortedDictionary
''''
''''        MsgBox ":"
''''        If Views Is Nothing Then
''''            MsgBox ":::"
''''
''''            Set Views = New_c.SortedDictionary(TextCompare)
''''
''''            ViewPath = App.Path & "\Resources\XAML\MainMenu.xml"
''''            With New XAMLReader
''''                Views.Add 0, .Load(New_c.FSO.ReadByteContent(ViewPath))
''''            End With
''''
''''            ViewPath = App.Path & "\Resources\XAML\SalesOrder.xml"
''''            With New XAMLReader
''''                Views.Add 1, .Load(New_c.FSO.ReadByteContent(ViewPath))
''''            End With
''''
''''        End If
        
120     Select Case Modulo
    
            Case 0: ViewPath = App.Path & "\Resources\XAML\MainMenu.xml"
122         Case 1: ViewPath = App.Path & "\Resources\XAML\SalesOrder.xml"
            'Case 2: ViewPath = ""
    
        End Select
                    
        New_c.Timing True
124     'Win.Widget.LockRefresh = True
126     Win.Children.Clear
        
128     With New XAMLReader
            Dim Cnt As Object
            If Len(ViewPath) Then Set Cnt = .Load(New_c.FSO.ReadByteContent(ViewPath))
            Debug.Print "XamlReader.Load(" & ViewPath & ")" & New_c.Timing
130         If Len(ViewPath) Then Win.Children.Add Cnt
            'Win.Children.Add Views(Modulo)
        End With
    
132     If Modulo = 1 Then
134         With New XAMLReader
136             LeftColumn.Children.Add .Load(New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\LeftColumn.xml"))
            End With
        End If
        
138     'Win.Widget.LockRefresh = False
        Debug.Print "Swap: " & New_c.Timing
        '<EhFooter>
        Exit Sub

SwapContent_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Test1.ShellWindow.SwapContent " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetContent()
On Error Resume Next

GoTo STEP1:

    Static bSet As Boolean
    If (Not bSet) And (Not RightColumn Is Nothing) Then
        bSet = True
        
        Form.SetFocus
        RightColumn.Widget.LockRefresh = True
        RightColumn.Children.Clear

        With New XAMLReader
            RightColumn.Children.Add .Load(New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\RightColumn.xml"))
        End With

        RightColumn.Widget.LockRefresh = False
        
        Exit Sub
    End If
    
    bSet = False
    
    'Debug.Print TypeName(Win.Form.ActiveWidget)
    
'    Static bExecuted As Boolean
'    If bExecuted Then Exit Sub
'    bExecuted = True
        
    Static ViewID As Long
    ViewID = ViewID + 1
    
    
    Dim ViewPath As String
    Dim Modulo As Long
    Modulo = ViewID Mod 3
    
    Select Case Modulo
    
        Case 0: ViewPath = App.Path & "\Resources\XAML\MainMenu.xml"
        Case 1: ViewPath = App.Path & "\Resources\XAML\SalesOrder.xml"
        Case 2: ViewPath = ""
    
    End Select
            
    Win.Widget.LockRefresh = True
    Win.Children.Clear
    
    With New XAMLReader
        If Len(ViewPath) Then Win.Children.Add .Load(New_c.FSO.ReadByteContent(ViewPath))
    End With
    
    If Modulo = 1 Then
        With New XAMLReader
            LeftColumn.Children.Add .Load(New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\LeftColumn.xml"))
        End With
    End If
        
    Win.Widget.LockRefresh = False
        
    Exit Sub
    
STEP1:
    Dim DC As SampleDataContext
    Set DC = Win.DataContext
    If DC Is Nothing Then Set Win.DataContext = Win.Children(0).DataContext
    
    Set DC = Win.DataContext
    If DC Is Nothing Then Exit Sub
    
    DC.Text = IIf(Not IsNumeric(DC.Text), 1, DC.Text)
    DC.Text = DC.Text * 1.1115 '1.005
    'DC.Text = IIf(DC.Text = 1, 1.5, 1)
    
    Exit Sub
    
    Static i As Long
    i = i + 1
    
    New_c.Timing True
    
    Dim UF As UniformGrid
    Dim ss As Scene
    Set ss = Win.Children(0)
    
    Dim v As Object
    
    For Each v In ss.Children
        If TypeOf v Is UniformGrid Then
            Set UF = v
            UF.Children.Clear
            UF.Children.Add New Button
            UF.Children.Add New Button
            UF.Children.Add New Button
            
            Dim Txb As TextBlock
            Set Txb = New TextBlock
            Txb.Text = "Hello"
            Set UF.Children(1) = Txb
            Exit For
        End If
    Next
'    Dim Img As Image
'    Set Img = New Image
'    Img.ImageKey = "menu.png"
'    Img.DesignLeft = 400
'    Img.DesignTop = 400
'    If i Mod 2 = 0 Then Img.DesignWidth = 100
'    Debug.Print i Mod 2 & ": " & Img.DesignWidth
'    Set m_Children(2) = Img
    
    Dim Btn As Button
    Set Btn = New Button
    Btn.DesignLeft = 300
    Btn.DesignTop = 300
    Btn.DesignWidth = 200
    Btn.DesignHeight = 100
    If i Mod 2 = 0 Then Btn.DesignWidth = 100
        
    Set Win.Children(0).Children.Item(2) = Btn
    Win.Form.Refresh
    
    Dim o As Object
    Dim Txt As String
    
    For Each o In Win.Children(0).Children
        If TypeOf o Is Panel Then
            Txt = IIf(i Mod 2 = 0, "Hello", "World")
            o.Children(0).Text = Txt
        End If
    Next
    'm_Children.Add Img

    Form.Caption = New_c.Timing
End Sub

'CSEH: ErrMsgBox
Private Sub InitializeComponent()
        '<EhHeader>
        On Error GoTo InitializeComponent_Err
        '</EhHeader>
        
        Win.DesignWidth = 1024
        Win.DesignHeight = 768
    
        Set Form = Win.Form
        Set NamedItems = Win.NamedChildren

100     Win.Form.WidgetRoot.LockRefresh = True
    
        'Win.Form.WidgetRoot.ImageKey = "menu.png"
        
    '''''''    Dim Obj As Button
    '''''''    Set Obj = New Button
    '''''''
    '''''''    Dim ImgA As VCF.Image
    '''''''    Set ImgA = New VCF.Image
    '''''''
    '''''''    Cairo.ImageList.AddImage "MyImg", "Resources\Close.png"
    '''''''    ImgA.ImageKey = "MyImg"
    '''''''
    '''''''
    '''''''    Obj.Children.Add ImgA
    '''''''
    '''''''    Dim tb As TextBlock
    '''''''    Set tb = New TextBlock
    '''''''    tb.Text = "Hello"
    '''''''    tb.DesignLeft = 60
    '''''''    Obj.Children.Add tb
    '''''''
    '''''''    Dim ss As VCF.Scene
    '''''''    Set ss = New VCF.Scene
    '''''''
    '''''''    Dim pnl1 As Panel
    '''''''    Set pnl1 = New Panel
    '''''''    ss.Children.Add pnl1
    '''''''    pnl1.Children.Add Obj
    '''''''
    '''''''    Win.Children.Add ss
    '''''''
    '''''''
    '''''''    MsgBox "1: " & TypeName(Obj.DataContext)
    '''''''    Set Win.DataContext = New Collection
    '''''''    MsgBox "2: " & TypeName(Obj.DataContext)
    '''''''    MsgBox "3: " & TypeName(ImgA.DataContext)
    '''''''    Set ImgA.DataContext = New ObservableCollection
    '''''''    MsgBox "4: " & TypeName(ImgA.DataContext)
    '''''''    MsgBox "5: " & TypeName(tb.DataContext)
    
102     With New XAMLReader
104         Win.Children.Add .Load(New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\MainMenu.xml"))
        End With
            
106     Win.Form.WidgetRoot.LockRefresh = False
    
108     'Set Win.DataContext = New Collection
        
        Exit Sub
    
        Dim Img As VCF.Image
    '    Set Img = New Image
    '    Img.ImageKey = "menu.png"
    '    Img.DesignLeft = 0
    '    Img.DesignTop = 0
    '    Img.DesignWidth = 800
    '    Img.DesignHeight = 600
    '    Call Win.Children.Add(Img)
    
    
        Dim b As Button
110     Set b = New Button
            
        Dim im As VCF.Image
112     Set im = New VCF.Image
114     With im
116         .ImageKey = "pic.png"
118         .DesignLeft = 75
120         .DesignTop = 25
122         .DesignWidth = 50
124         .DesignHeight = 50
        End With
    
    
        Dim TT As TextBlock
126     Set TT = New TextBlock
128     With TT
130         .DesignTop = 70
132         .DesignLeft = 50
134         .Text = "Caption"
        End With
    
136     With b
138         .DesignWidth = 200
140         .DesignHeight = 100
142         .DesignLeft = 8
144         .DesignTop = 8
146         .Children.Add im
148         .Children.Add TT
        End With
    
150     Call Win.Children.Add(b)
    
152     Set b = New Button
154     With b
156         .DesignLeft = 216
158         .DesignTop = 8
160         .DesignWidth = 200
162         .DesignHeight = 100
        
164         .Children.Add New Button
166         With .Children(0)
168             .DesignWidth = 184
170             .DesignHeight = 84
172             .DesignLeft = 8
174             .DesignTop = 8
            End With
        End With
    
176     Call Win.Children.Add(b)
    
    
178     Set b = New Button
180     b.DesignLeft = 424
182     Call Win.Children.Add(b)
    
184     Set b = New Button
186     b.DesignLeft = 632
188     Call Win.Children.Add(b)
    
190     Set b = New Button
192     b.DesignTop = 116
194     b.DesignLeft = 8
196     Call Win.Children.Add(b)
    
198     Set b = New Button
200     b.DesignTop = 116
202     b.DesignLeft = 216
204     Call Win.Children.Add(b)
    
206     Set Img = New VCF.Image
208     Img.ImageKey = "pic.png"
210     Img.DesignTop = 350
212     Img.DesignHeight = 150
214     Img.DesignWidth = 250
216     Call Win.Children.Add(Img)
    
        Dim Txt As TextBlock
218     Set Txt = New TextBlock
220     Txt.DesignTop = 8
222     Txt.DesignLeft = 8
224     Txt.DesignHeight = 30
226     Txt.DesignWidth = 100
    
        Dim pnl As Panel
228     Set pnl = New Panel
230     With pnl
232         .DesignLeft = 300
234         .DesignTop = 350
236         .Children.Add Txt
        End With
    
238     Call Win.Children.Add(pnl)
        '<EhFooter>
        Exit Sub

InitializeComponent_Err:
        Debug.Print Err.Description & vbCrLf & _
               "in Test1.ShellWindow.InitializeComponent " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Click()
    SetContent
    
    If Not NamedItems.ContainsKey("OrderItems") Then Exit Sub
    
    Dim LV As ListView
    Set LV = NamedItems("OrderItems")
    
    Dim oc As ObservableCollection
    Set oc = LV.ItemsSource
    
    Dim v
    Dim Item As ObservableItem
    
    If oc Is Nothing Then Exit Sub
    
    For Each v In oc
        If IsObject(v) Then
            If TypeOf v Is ObservableItem Then
                Set Item = v
                Exit For
            End If
        End If
    Next
    
    If Item Is Nothing Then Exit Sub
    
    Item.FirstName = IIf(Item.FirstName = "Doros", "John", "Doros")
    Item.LastName = IIf(Item.LastName = "Dimitriou", "Smith", "Dimitriou")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Try:
    On Error GoTo Catch
    
    Static IsBusy As Boolean
    Static StopSignal As Boolean
    
    Dim Index As Long
    
    If KeyCode <> 13 Then Exit Sub
    
    If IsBusy Then
        IsBusy = False
        StopSignal = True
        Exit Sub
    End If
    
    IsBusy = True
    StopSignal = False
    
    New_c.Timing True
    
    For Index = 0 To 49999
        If StopSignal Then Exit For
        Form.Caption = Index + 1 & " in " & New_c.Timing
        SetContent
        SetContent
        SetContent
        DoEvents
    Next

    Exit Sub
Catch:
    '
End Sub

Private Sub Form_Load()
    'Form.WidgetRoot.DesignMode = True
    'Form.BorderStyle = vbBSNone
    
    Form.CenterOn New_c.Displays(1)
    'Form.FullScreen = True
End Sub

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = Win
End Property

Private Sub IWindow_InitializeComponent()
    InitializeComponent
End Sub

Private Sub NamedItems_CollectionChanged(sender As Object, e As VCF.CollectionChangedEventArgs)
    Set RightColumn = Nothing
    Set LeftColumn = Nothing
    Set MyHost = Nothing
    
    If NamedItems.ContainsKey("RightColumn") Then Set RightColumn = NamedItems("RightColumn")
    If NamedItems.ContainsKey("LeftColumn") Then Set LeftColumn = NamedItems("LeftColumn")
    
    If NamedItems.ContainsKey("OrderItems") Then
        Dim LV As ListView
        Set LV = NamedItems("OrderItems")
        If Not LV.ItemsSource Is Nothing Then Exit Sub
        
'        With New XAMLReader
'            Set LV.ItemTemplate = .Load(New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\DataTemplate1.xml"))
'        End With
        
'''        Dim oc As ObservableCollection
'''        Set oc = New ObservableCollection
'''        oc.Add "Hello"
'''        oc.Add "World"
'''
'''        Dim Obj As ObservableItem
'''        Set Obj = New ObservableItem
'''        Obj.FirstName = "Doros"
'''        Obj.LastName = "Dimitriou"
'''
'''        oc.Add Obj
'''
'''        Set Obj = New ObservableItem
'''        Obj.FirstName = "Hasan"
'''        Obj.LastName = "Poullis"
'''
'''        oc.Add Obj
'''
'''        oc.Add New SampleListItem
'''
'''        Set LV.ItemsSource = oc
    End If
End Sub
