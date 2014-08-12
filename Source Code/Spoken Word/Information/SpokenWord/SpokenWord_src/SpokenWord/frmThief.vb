Public Class frmThief
    Inherits System.Windows.Forms.Form

    ' This window does NOTHING except steal focus from 
    ' the currently active application for a split second.
    ' Why do we do this?  WORD is not accessible to the
    ' app until it has lost focus.  Why?  Who knows!
    ' In any case, if you disable this window, half
    ' the time the Word doc will be invisible to the 
    ' "hook" process!

    Private _ticks As Integer = 0
    Private _oldFocus As IntPtr = IntPtr.Zero

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Timer1.Enabled = True
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container
Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
'
'Timer1
'
Me.Timer1.Interval = 100
'
'frmThief
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(104, 5)
Me.ControlBox = False
Me.MaximizeBox = False
Me.MinimizeBox = False
Me.Name = "frmThief"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "Seeking Word..."

    End Sub

#End Region

    Private Declare Function GetCurrentThreadId Lib "kernel32" () As IntPtr

    Private Declare Function AttachThreadInput Lib "user32" _
        (ByVal idAttach As IntPtr, _
         ByVal idAttachTo As IntPtr, _
         ByVal fAttach As Boolean) As Boolean

    Private Declare Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hWnd As IntPtr, _
         ByVal lpdwProcessId As IntPtr) As IntPtr

    Private Declare Function SetForegroundWindow Lib "user32" (ByVal ByValhWnd As IntPtr) As Boolean
    Private Declare Function GetForegroundWindow Lib "user32" () As IntPtr

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As _
            System.EventArgs) Handles Timer1.Tick

        _ticks += Timer1.Interval

        If _oldFocus.Equals(IntPtr.Zero) Then
            ' save old focus.
            _oldFocus = GetForegroundWindow()
        End If

        Dim foregroundThread As IntPtr = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero)
        'Console.WriteLine("ForegroundThread = " & foregroundThread.ToString)
        Dim currentThread As IntPtr = GetCurrentThreadId()
        'Console.WriteLine("CurrentThread = " & currentThread.ToString)

        If _ticks >= 500 Then
            Timer1.Enabled = False
            Try
                ' set focus back.  Doesn't seem to be working.  DO NOT TRACE INTO!!
                currentThread = GetWindowThreadProcessId(_oldFocus, IntPtr.Zero)
                If Not foregroundThread.Equals(currentThread) Then
                    AttachThreadInput(foregroundThread, currentThread, True)
                End If
                SetForegroundWindow(_oldFocus)
                If Not foregroundThread.Equals(currentThread) Then
                    'Console.WriteLine("Detaching input")
                    AttachThreadInput(foregroundThread, currentThread, False)
                End If
            Catch
            End Try
            Application.DoEvents()
            Me.Close()
        End If

        If Me.WindowState = FormWindowState.Minimized Then
            Me.WindowState = FormWindowState.Normal
        End If

        If Not foregroundThread.Equals(currentThread) Then
            'Console.WriteLine("Attaching input")
            AttachThreadInput(foregroundThread, currentThread, True)
        Else
            ' we have focus.  Terminate loop next pass.
            _ticks = 500
        End If

        SetForegroundWindow(Me.Handle)

        If Not foregroundThread.Equals(currentThread) Then
            'Console.WriteLine("Detaching input")
            AttachThreadInput(foregroundThread, currentThread, False)
        End If
    End Sub

End Class
