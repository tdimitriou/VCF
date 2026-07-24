VERSION 1.0 CLASS

BEGIN

  MultiUse = -1  'True

  Persistable = 0  'NotPersistable

  DataBindingBehavior = 0  'vbNone

  DataSourceBehavior  = 0  'vbNone

  MTSTransactionMode  = 0  'NotAnMTSObject

END

Attribute VB_Name = "BorderTestWindow"

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

Private WithEvents m_ChromeDiagTimer As cTimer

Attribute m_ChromeDiagTimer.VB_VarHelpID = -1



Private m_DiagPollTicks As Long

Private m_DiagLastFormBorder As Long

Private m_DiagFormBorderBeforeShow As Long

Private m_DiagChangeCountAfterShow As Long

Private m_DiagLoggedActivate As Boolean



Public Sub Show(Optional Modal As FormShowConstants = vbModeless)

    m_DiagFormBorderBeforeShow = Form.BorderStyle

    modBorderChromeDiag.LogStage "Show BEFORE Win.Show", Win, Form

    

    Win.Show Modal

    

    modBorderChromeDiag.LogStage "Show AFTER Win.Show", Win, Form

    BeginChromePoll

End Sub



Private Sub Class_Initialize()

    NewWindow Me, Win

    Set Form = Win.Form

    modBorderChromeDiag.LogStage "After NewWindow", Win, Form

End Sub



Private Sub IWindow_InitializeComponent()

    Dim Xml As String
    Dim Started As Double

    

    If Not modApp.XamlResources.ContainsKey("BorderTestWindow") Then

        Err.Raise vbObjectError, , "BorderTestWindow.xml missing from Resources\XAML"

    End If

    

    Xml = modApp.XamlResources("BorderTestWindow")

    Started = Timer

    With New XAMLReader
        .LoadSuperclassData Me, Xml
    End With

    modHarnessLoadBench.LogTimedStage "P7d-LOAD-BORDERED", "LoadSuperclassData", modHarnessLoadBench.ElapsedMsSince(Started)

    modBorderChromeDiag.LogStage "After LoadSuperclassData (XAML BorderStyle=2)", Win, Win.Form
End Sub



Private Property Get IWindow_Base() As VCF.Window

    Set IWindow_Base = Win

End Property



Private Sub Form_Load()

    modBorderChromeDiag.LogStage "Form_Load", Win, Form

End Sub



Private Sub Form_Activate()

    If m_DiagLoggedActivate Then Exit Sub

    m_DiagLoggedActivate = True

    modBorderChromeDiag.LogStage "Form_Activate (first)", Win, Form

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Form.Unload

End Sub



Private Sub Form_QueryUnload(Cancel As Boolean, UnloadMode As QueryUnloadConstants)

    StopChromePoll

End Sub



Private Sub BeginChromePoll()

    m_DiagPollTicks = 0

    m_DiagChangeCountAfterShow = 0

    m_DiagLastFormBorder = Form.BorderStyle

    modBorderChromeDiag.LogStage "Poll start (50ms x 20)", Win, Form

    Set m_ChromeDiagTimer = New_c.Timer(50, True)

End Sub



Private Sub StopChromePoll()

    On Error Resume Next

    Set m_ChromeDiagTimer = Nothing

End Sub



Private Sub m_ChromeDiagTimer_Timer()

    Dim Current As Long

    Current = Form.BorderStyle

    

    If m_DiagPollTicks = 0 Then

        If Current <> m_DiagLastFormBorder Then

            modBorderChromeDiag.LogChange "first poll tick", m_DiagLastFormBorder, Current

            m_DiagChangeCountAfterShow = m_DiagChangeCountAfterShow + 1

        End If

    ElseIf Current <> m_DiagLastFormBorder Then

        modBorderChromeDiag.LogChange "poll tick " & m_DiagPollTicks, m_DiagLastFormBorder, Current

        m_DiagChangeCountAfterShow = m_DiagChangeCountAfterShow + 1

    End If

    

    m_DiagLastFormBorder = Current

    m_DiagPollTicks = m_DiagPollTicks + 1

    

    If m_DiagPollTicks >= 20 Then

        StopChromePoll

        modBorderChromeDiag.LogSummary m_DiagFormBorderBeforeShow, Current, m_DiagChangeCountAfterShow

    End If

End Sub


