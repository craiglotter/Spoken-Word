Public Class frmHover
    Inherits System.Windows.Forms.Form

Private WithEvents m_voice As VoiceHelper
Private WithEvents m_Word As WordHelper
Private m_IsReading As Boolean = False

Private Const WM_NCLBUTTONDOWN As Integer = &HA1
Private Const HTCAPTION As Integer = 2

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.ClientSize = New Size(tb_FINAL.Rectangle.X + 3, Me.ToolBar1.Height + Me.sb.Height)
        cboEffect.SelectedIndex = 0
        m_Word = New WordHelper
        FindDoc()
        Try
            m_voice = New VoiceHelper
            Me.hsSpeed.Value = m_voice.Speed
            Me.hsPitch.Value = m_voice.Pitch
            cboVoice.Items.AddRange(m_voice.Voices.ToArray)
            cboVoice.Text = m_voice.CurrentVoice
        Catch ex As Exception
            ShowStatus("No Speech installed!")
            tbAcquire.Enabled = False
            tbRead.Enabled = False
        End Try
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
    Friend WithEvents ToolBar1 As System.Windows.Forms.ToolBar
    Friend WithEvents tbAcquire As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbRead As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbPause As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbBack As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbSkip As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbRepeat As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbStop As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tb_FINAL As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbExpand As System.Windows.Forms.ToolBarButton
    Friend WithEvents hsPitch As System.Windows.Forms.HScrollBar
    Friend WithEvents hsSpeed As System.Windows.Forms.HScrollBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboVoice As System.Windows.Forms.ComboBox
    Friend WithEvents cboEffect As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents sb As System.Windows.Forms.StatusBar
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tbRaw As System.Windows.Forms.TabPage
    Friend WithEvents tbPre As System.Windows.Forms.TabPage
    Friend WithEvents tbEffect As System.Windows.Forms.TabPage
    Friend WithEvents tbPost As System.Windows.Forms.TabPage
    Friend WithEvents tbSAPI As System.Windows.Forms.TabPage
    Friend WithEvents rtbOriginal As System.Windows.Forms.RichTextBox
    Friend WithEvents txtPreformat As System.Windows.Forms.TextBox
    Friend WithEvents tbLog As System.Windows.Forms.TabPage
    Friend WithEvents txtEffect As System.Windows.Forms.TextBox
    Friend WithEvents txtPostformat As System.Windows.Forms.TextBox
    Friend WithEvents txtSAPI As System.Windows.Forms.TextBox
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHover))
        Me.ToolBar1 = New System.Windows.Forms.ToolBar
        Me.tbAcquire = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
        Me.tbRead = New System.Windows.Forms.ToolBarButton
        Me.tbPause = New System.Windows.Forms.ToolBarButton
        Me.tbStop = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton
        Me.tbBack = New System.Windows.Forms.ToolBarButton
        Me.tbRepeat = New System.Windows.Forms.ToolBarButton
        Me.tbSkip = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton3 = New System.Windows.Forms.ToolBarButton
        Me.tbExpand = New System.Windows.Forms.ToolBarButton
        Me.tb_FINAL = New System.Windows.Forms.ToolBarButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.hsPitch = New System.Windows.Forms.HScrollBar
        Me.hsSpeed = New System.Windows.Forms.HScrollBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboVoice = New System.Windows.Forms.ComboBox
        Me.cboEffect = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.sb = New System.Windows.Forms.StatusBar
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tbRaw = New System.Windows.Forms.TabPage
        Me.rtbOriginal = New System.Windows.Forms.RichTextBox
        Me.tbPre = New System.Windows.Forms.TabPage
        Me.txtPreformat = New System.Windows.Forms.TextBox
        Me.tbEffect = New System.Windows.Forms.TabPage
        Me.txtEffect = New System.Windows.Forms.TextBox
        Me.tbPost = New System.Windows.Forms.TabPage
        Me.txtPostformat = New System.Windows.Forms.TextBox
        Me.tbSAPI = New System.Windows.Forms.TabPage
        Me.txtSAPI = New System.Windows.Forms.TextBox
        Me.tbLog = New System.Windows.Forms.TabPage
        Me.txtLog = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TabControl1.SuspendLayout()
        Me.tbRaw.SuspendLayout()
        Me.tbPre.SuspendLayout()
        Me.tbEffect.SuspendLayout()
        Me.tbPost.SuspendLayout()
        Me.tbSAPI.SuspendLayout()
        Me.tbLog.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolBar1
        '
        Me.ToolBar1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tbAcquire, Me.ToolBarButton1, Me.tbRead, Me.tbPause, Me.tbStop, Me.ToolBarButton2, Me.tbBack, Me.tbRepeat, Me.tbSkip, Me.ToolBarButton3, Me.tbExpand, Me.tb_FINAL})
        Me.ToolBar1.DropDownArrows = True
        Me.ToolBar1.ImageList = Me.ImageList1
        Me.ToolBar1.Location = New System.Drawing.Point(0, 0)
        Me.ToolBar1.Name = "ToolBar1"
        Me.ToolBar1.ShowToolTips = True
        Me.ToolBar1.Size = New System.Drawing.Size(306, 28)
        Me.ToolBar1.TabIndex = 0
        '
        'tbAcquire
        '
        Me.tbAcquire.ImageIndex = 0
        Me.tbAcquire.Tag = "HOOK"
        Me.tbAcquire.ToolTipText = "Open a Word doc and click to begin."
        '
        'ToolBarButton1
        '
        Me.ToolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbRead
        '
        Me.tbRead.ImageIndex = 1
        Me.tbRead.Tag = "READ"
        Me.tbRead.ToolTipText = "Click to begin reading highlighted text (or from cursor)"
        '
        'tbPause
        '
        Me.tbPause.ImageIndex = 2
        Me.tbPause.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbPause.Tag = "PAUSE_RESUME"
        Me.tbPause.ToolTipText = "Click to pause/resume playback"
        '
        'tbStop
        '
        Me.tbStop.ImageIndex = 3
        Me.tbStop.Tag = "STOP"
        Me.tbStop.ToolTipText = "Click to stop playback"
        '
        'ToolBarButton2
        '
        Me.ToolBarButton2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbBack
        '
        Me.tbBack.ImageIndex = 4
        Me.tbBack.Tag = "BACK"
        Me.tbBack.ToolTipText = "Click to read the previous paragraph"
        '
        'tbRepeat
        '
        Me.tbRepeat.ImageIndex = 5
        Me.tbRepeat.Tag = "REPEAT"
        Me.tbRepeat.ToolTipText = "Click to repeat the current selection"
        '
        'tbSkip
        '
        Me.tbSkip.ImageIndex = 6
        Me.tbSkip.Tag = "SKIP"
        Me.tbSkip.ToolTipText = "Click to skip to the next paragraph"
        '
        'ToolBarButton3
        '
        Me.ToolBarButton3.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbExpand
        '
        Me.tbExpand.ImageIndex = 7
        Me.tbExpand.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbExpand.Tag = "EXPAND_COLLAPSE"
        Me.tbExpand.ToolTipText = "Click for additional options"
        '
        'tb_FINAL
        '
        Me.tb_FINAL.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Silver
        '
        'hsPitch
        '
        Me.hsPitch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.hsPitch.LargeChange = 5
        Me.hsPitch.Location = New System.Drawing.Point(48, 88)
        Me.hsPitch.Maximum = 15
        Me.hsPitch.Minimum = -10
        Me.hsPitch.Name = "hsPitch"
        Me.hsPitch.Size = New System.Drawing.Size(250, 16)
        Me.hsPitch.TabIndex = 12
        '
        'hsSpeed
        '
        Me.hsSpeed.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.hsSpeed.LargeChange = 5
        Me.hsSpeed.Location = New System.Drawing.Point(48, 64)
        Me.hsSpeed.Maximum = 15
        Me.hsSpeed.Minimum = -10
        Me.hsSpeed.Name = "hsSpeed"
        Me.hsSpeed.Size = New System.Drawing.Size(250, 16)
        Me.hsSpeed.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Speed"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(0, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Pitch"
        '
        'cboVoice
        '
        Me.cboVoice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboVoice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboVoice.Location = New System.Drawing.Point(48, 32)
        Me.cboVoice.Name = "cboVoice"
        Me.cboVoice.Size = New System.Drawing.Size(248, 21)
        Me.cboVoice.TabIndex = 15
        '
        'cboEffect
        '
        Me.cboEffect.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboEffect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEffect.Items.AddRange(New Object() {"(None)"})
        Me.cboEffect.Location = New System.Drawing.Point(48, 112)
        Me.cboEffect.Name = "cboEffect"
        Me.cboEffect.Size = New System.Drawing.Size(250, 21)
        Me.cboEffect.TabIndex = 16
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(0, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Effect"
        '
        'sb
        '
        Me.sb.Location = New System.Drawing.Point(0, 428)
        Me.sb.Name = "sb"
        Me.sb.Size = New System.Drawing.Size(306, 22)
        Me.sb.SizingGrip = False
        Me.sb.TabIndex = 18
        Me.sb.Text = "initializing..."
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.tbRaw)
        Me.TabControl1.Controls.Add(Me.tbPre)
        Me.TabControl1.Controls.Add(Me.tbEffect)
        Me.TabControl1.Controls.Add(Me.tbPost)
        Me.TabControl1.Controls.Add(Me.tbSAPI)
        Me.TabControl1.Controls.Add(Me.tbLog)
        Me.TabControl1.Location = New System.Drawing.Point(0, 136)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(306, 288)
        Me.TabControl1.TabIndex = 19
        '
        'tbRaw
        '
        Me.tbRaw.BackColor = System.Drawing.Color.LightBlue
        Me.tbRaw.Controls.Add(Me.rtbOriginal)
        Me.tbRaw.Location = New System.Drawing.Point(4, 22)
        Me.tbRaw.Name = "tbRaw"
        Me.tbRaw.Size = New System.Drawing.Size(298, 262)
        Me.tbRaw.TabIndex = 0
        Me.tbRaw.Text = "Orig"
        '
        'rtbOriginal
        '
        Me.rtbOriginal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbOriginal.Location = New System.Drawing.Point(0, 0)
        Me.rtbOriginal.Name = "rtbOriginal"
        Me.rtbOriginal.ReadOnly = True
        Me.rtbOriginal.Size = New System.Drawing.Size(298, 262)
        Me.rtbOriginal.TabIndex = 0
        Me.rtbOriginal.Text = ""
        '
        'tbPre
        '
        Me.tbPre.BackColor = System.Drawing.Color.LightBlue
        Me.tbPre.Controls.Add(Me.txtPreformat)
        Me.tbPre.Location = New System.Drawing.Point(4, 22)
        Me.tbPre.Name = "tbPre"
        Me.tbPre.Size = New System.Drawing.Size(298, 262)
        Me.tbPre.TabIndex = 1
        Me.tbPre.Text = "Pre"
        '
        'txtPreformat
        '
        Me.txtPreformat.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtPreformat.Location = New System.Drawing.Point(0, 0)
        Me.txtPreformat.Multiline = True
        Me.txtPreformat.Name = "txtPreformat"
        Me.txtPreformat.ReadOnly = True
        Me.txtPreformat.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtPreformat.Size = New System.Drawing.Size(298, 262)
        Me.txtPreformat.TabIndex = 0
        Me.txtPreformat.Text = ""
        '
        'tbEffect
        '
        Me.tbEffect.BackColor = System.Drawing.Color.LightBlue
        Me.tbEffect.Controls.Add(Me.txtEffect)
        Me.tbEffect.Location = New System.Drawing.Point(4, 22)
        Me.tbEffect.Name = "tbEffect"
        Me.tbEffect.Size = New System.Drawing.Size(298, 262)
        Me.tbEffect.TabIndex = 2
        Me.tbEffect.Text = "Effect"
        '
        'txtEffect
        '
        Me.txtEffect.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtEffect.Location = New System.Drawing.Point(0, 0)
        Me.txtEffect.Multiline = True
        Me.txtEffect.Name = "txtEffect"
        Me.txtEffect.ReadOnly = True
        Me.txtEffect.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEffect.Size = New System.Drawing.Size(298, 262)
        Me.txtEffect.TabIndex = 1
        Me.txtEffect.Text = ""
        '
        'tbPost
        '
        Me.tbPost.BackColor = System.Drawing.Color.LightBlue
        Me.tbPost.Controls.Add(Me.txtPostformat)
        Me.tbPost.Location = New System.Drawing.Point(4, 22)
        Me.tbPost.Name = "tbPost"
        Me.tbPost.Size = New System.Drawing.Size(298, 262)
        Me.tbPost.TabIndex = 3
        Me.tbPost.Text = "Post"
        '
        'txtPostformat
        '
        Me.txtPostformat.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtPostformat.Location = New System.Drawing.Point(0, 0)
        Me.txtPostformat.Multiline = True
        Me.txtPostformat.Name = "txtPostformat"
        Me.txtPostformat.ReadOnly = True
        Me.txtPostformat.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtPostformat.Size = New System.Drawing.Size(298, 262)
        Me.txtPostformat.TabIndex = 1
        Me.txtPostformat.Text = ""
        '
        'tbSAPI
        '
        Me.tbSAPI.BackColor = System.Drawing.Color.LightBlue
        Me.tbSAPI.Controls.Add(Me.txtSAPI)
        Me.tbSAPI.Location = New System.Drawing.Point(4, 22)
        Me.tbSAPI.Name = "tbSAPI"
        Me.tbSAPI.Size = New System.Drawing.Size(298, 262)
        Me.tbSAPI.TabIndex = 4
        Me.tbSAPI.Text = "SAPI"
        '
        'txtSAPI
        '
        Me.txtSAPI.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSAPI.Location = New System.Drawing.Point(0, 0)
        Me.txtSAPI.Multiline = True
        Me.txtSAPI.Name = "txtSAPI"
        Me.txtSAPI.ReadOnly = True
        Me.txtSAPI.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSAPI.Size = New System.Drawing.Size(298, 262)
        Me.txtSAPI.TabIndex = 1
        Me.txtSAPI.Text = ""
        '
        'tbLog
        '
        Me.tbLog.BackColor = System.Drawing.Color.LightBlue
        Me.tbLog.Controls.Add(Me.txtLog)
        Me.tbLog.Location = New System.Drawing.Point(4, 22)
        Me.tbLog.Name = "tbLog"
        Me.tbLog.Size = New System.Drawing.Size(298, 262)
        Me.tbLog.TabIndex = 5
        Me.tbLog.Text = "Log"
        '
        'txtLog
        '
        Me.txtLog.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtLog.Location = New System.Drawing.Point(0, 0)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ReadOnly = True
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtLog.Size = New System.Drawing.Size(298, 262)
        Me.txtLog.TabIndex = 1
        Me.txtLog.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(0, 36)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 16)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Voice"
        '
        'frmHover
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightBlue
        Me.ClientSize = New System.Drawing.Size(306, 450)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.sb)
        Me.Controls.Add(Me.cboEffect)
        Me.Controls.Add(Me.cboVoice)
        Me.Controls.Add(Me.hsPitch)
        Me.Controls.Add(Me.hsSpeed)
        Me.Controls.Add(Me.ToolBar1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(312, 482)
        Me.Name = "frmHover"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Spoken Word"
        Me.TabControl1.ResumeLayout(False)
        Me.tbRaw.ResumeLayout(False)
        Me.tbPre.ResumeLayout(False)
        Me.tbEffect.ResumeLayout(False)
        Me.tbPost.ResumeLayout(False)
        Me.tbSAPI.ResumeLayout(False)
        Me.tbLog.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Function FindDoc() As Boolean
        Me.tbBack.Enabled = False
        Me.tbPause.Enabled = False
        Me.tbRead.Enabled = False
        Me.tbRepeat.Enabled = False
        Me.tbSkip.Enabled = False
        Me.tbStop.Enabled = False
        Me.tbPause.Pushed = False

        Dim doc As Word.Document
        doc = m_Word.CurrentDocument
        If doc Is Nothing Then
            Me.Text = "Spoken Word: [no document found]"
            ShowStatus("Open a Word doc, then click [book] to begin.")
            Return False
        Else
            Me.Text = "Spoken Word: " & doc.Name
            ShowStatus("Ready.")
            Me.tbRead.Enabled = True
            Return True
        End If
    End Function

    Private Sub ReadSelection()
        m_voice.Pitch = hsPitch.Value
        m_voice.Speed = hsSpeed.Value
        m_voice.SetVoice(cboVoice.SelectedItem.ToString)

        Dim sel As Word.Selection = m_Word.GetSelection()
        Dim txt As String = FixText(sel.Text)
        If m_Word.IsAnIndexBlock(sel) Then txt = "(Index or Table of Contents)"
        rtbOriginal.Clear()
        If txt > "" Then
            Try
                sel.Copy()
            Catch
                Clipboard.SetDataObject("")
            End Try
            Dim clip As IDataObject = Clipboard.GetDataObject
            If clip.GetDataPresent(GetType(String)) Then
                ' richtext control won't paste while readonly!
                rtbOriginal.ReadOnly = False
                rtbOriginal.Paste()     ' RTB format (if avaliable), but numbered lists will be wrong.
                rtbOriginal.ReadOnly = True
                txtPreformat.Text = clip.GetData(GetType(String)).ToString  ' this will return the correct numbering.
            End If

            txtPreformat.Text = PreProcess(FixText(txtPreformat.Text))
            ApplyEffect()
            txtPostformat.Text = PostProcess(txtEffect.Text)
            txtSAPI.Text = m_voice.SAPIfy(txtPostformat.Text)
        Else
            rtbOriginal.Text = "(Unknown content)"
            txtPreformat.Text = ""
            txtEffect.Text = ""
            txtPostformat.Text = ""
            txtSAPI.Text = "<silence msec=""250""/>"
        End If

        m_voice.Speak(txtSAPI.Text, sel.End)
    End Sub

    Private Sub ApplyEffect()
        txtEffect.Text = txtPreformat.Text
        If cboEffect.SelectedIndex > 0 Then
            Dim effect As String = cboEffect.SelectedItem.ToString
            If effect = "(Random)" Then
                Dim i As Integer = CInt(Int((cboEffect.Items.Count - 2) * Rnd() + 1))
                Console.WriteLine("RANDOM EFFECT: " & cboEffect.Items(i).ToString)
            End If
            ShowStatus("Applying effect: " & effect)
            Dim parts As String() = effect.Split(":"c)
            Dim tmp As String = txtPreformat.Text
            Try
                Select Case parts(0)
                    Case "Borland"
                        Dim borland As New com.borland.ww6.IBorlandBabelservice
                        borland.Timeout = 1000
                        tmp = borland.BabelFish(Trim(parts(1)), tmp)
                    Case "AspxRunway"
                        Dim aspxrunway As New com.aspxpressway.www.piglatin
                        aspxrunway.Timeout = 1000
                        tmp = aspxrunway.toPigLatin(tmp)
                    Case "BabelFish"
                        Dim babelfish As New net.xmethods.www.BabelFishService
                        babelfish.Timeout = 2000 ' translation services are kinda slow
                        ShowStatus("... translating to French")
                        tmp = babelfish.BabelFish("en_fr", tmp)
                        ShowStatus("... translating to German")
                        tmp = babelfish.BabelFish("fr_de", tmp)
                        ShowStatus("... translating to English")
                        tmp = babelfish.BabelFish("de_en", tmp)
                    Case "WebserviceX"
                        Dim webx As New net.webservicex.www.TranslationService
                        Dim lang As net.webservicex.www.Language
                        webx.Timeout = 2000 ' translation services are kinda slow
                        ShowStatus("... translating to French")
                        tmp = webx.Translate(lang.EnglishTOFrench, tmp)
                        ShowStatus("... translating to German")
                        tmp = webx.Translate(lang.FrenchTOGerman, tmp)
                        ShowStatus("... translating to English")
                        tmp = webx.Translate(lang.GermanTOEnglish, tmp)
                End Select
                ' clean up response.
                txtEffect.Text = FixText(Decode(tmp))
            Catch ex As Exception
                ShowStatus("ERROR calling webservice: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub m_voice_Bookmark(ByVal Bookmark As String, ByVal BookmarkId As Integer) Handles m_voice.Bookmark
        ' currently get bookmark = -1 at START of text
        ' bookmark of (end position) at END of text
        Try
            Dim test As Integer = Integer.Parse(Bookmark)
        Catch
            ' Do nothing (errant bookmark).  for some reason, currency values are being raised as bookmarks!!!!!
            ' may be other errant bookmarks -- this should weed out most of them.
            ShowStatus("Errant bookmark encountered: " & Bookmark)
            Exit Sub
        End Try

        ' The idea here was that the bookmark would contain positional information
        ' so that as stuff was read, it could be synchronized with the Word doc.
        ' However, I never got this working correctly, and just "letting it go"
        ' caused fewer errors (syncronization tended to get really lost and
        ' start looping inside tables and such).
        If BookmarkId < 0 Then
            'Console.WriteLine("CURRENT: " & m_Word.GetSelection().Text)
            'Console.WriteLine("NEXT: " & m_Word.PeekNext().Text)

            'ShowStatus("Synchronization: MoveTo " & BookmarkId)
            'm_Word.MoveTo(BookmarkId)
        ElseIf Me.tbStop.Enabled Then
            ' synchronization cue.  Isn't working correctly!!
            'm_Word.SetCaret(BookmarkId)

            ' feed in next block of text.
            If m_Word.NavNext() Then
                ReadSelection()
            Else
                ShowStatus("Bookmark: No further text to read.")
                Try
                    m_Word.SetCaret(m_Word.GetSelection.Range.End)
                Catch
                End Try
                Me.tbBack.Enabled = False
                Me.tbPause.Enabled = False
                Me.tbRepeat.Enabled = False
                Me.tbSkip.Enabled = False
                Me.tbStop.Enabled = False
                Me.tbPause.Pushed = False
            End If
        End If
    End Sub

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick
        Select Case e.Button.Tag
            Case "HOOK"
                ' try without the thief first...
                If Not FindDoc() Then
                    ' this helps locate the word engine.  don't ask me why!
                    Dim pop As New frmThief
                    pop.ShowDialog()
                    Application.DoEvents() ' wait for it....
                    FindDoc()
                End If
            Case "READ"
                m_voice.StopSpeaking()
                Dim sel As Word.Selection = m_Word.GetSelection()
                If sel Is Nothing Then Exit Sub
                m_Word.SetLimits() ' remember bounds
                sel.SetRange(sel.Range.Start, sel.Range.Start) ' move to start
                m_Word.NavNext()
                Me.tbBack.Enabled = True
                Me.tbPause.Enabled = True
                Me.tbPause.Pushed = False
                Me.tbRead.Enabled = True
                Me.tbRepeat.Enabled = True
                Me.tbSkip.Enabled = True
                Me.tbStop.Enabled = True
                ReadSelection()
            Case "STOP"
                Me.tbBack.Enabled = False
                Me.tbPause.Enabled = False
                Me.tbRepeat.Enabled = False
                Me.tbSkip.Enabled = False
                Me.tbStop.Enabled = False
                Me.tbPause.Pushed = False
                m_voice.StopSpeaking()
                Try
                    m_Word.SetCaret(m_Word.GetSelection.Range.End)
                Catch
                End Try
            Case "PAUSE_RESUME"
                If e.Button.Pushed Then
                    m_voice.PauseSpeaking()
                Else
                    m_voice.ResumeSpeaking()
                End If
            Case "BACK"
                ' repeat prior paragraph/section
                m_voice.StopSpeaking()
                If m_Word.NavPrior() Then ReadSelection() Else ShowStatus("Back: No prior text to read.")
            Case "REPEAT"
                ' repeat current sentence S-L-O-W-E-R
                m_voice.Repeat()
            Case "SKIP"
                ' skip to next paragraph/section
                m_voice.StopSpeaking()
                If m_Word.NavNext() Then ReadSelection() Else ShowStatus("Skip: No further text to read.")
            Case "EXPAND_COLLAPSE"
                If Me.tbExpand.Pushed Then
                    Me.ClientSize = New Size(270, 470)
                Else
                    Me.ClientSize = New Size(tb_FINAL.Rectangle.X + 3, Me.ToolBar1.Height + Me.sb.Height)
                End If
            Case Else
                Console.WriteLine("case """ & e.Button.Tag.ToString & """")
        End Select
    End Sub

    Private Sub cboVoice_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboVoice.SelectedIndexChanged
        ShowStatus("Voice changed.  Will take effect at next transition.")
    End Sub

    Private Sub hsSpeed_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles hsSpeed.ValueChanged
        ShowStatus("Speed adjusted.  Will take effect at next transition.")
    End Sub

    Private Sub hsPitch_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles hsPitch.ValueChanged
        ShowStatus("Pitch adjusted.  Will take effect at next transition.")
    End Sub

    Private Sub ShowStatus(ByVal Msg As String) Handles m_voice.Status, m_Word.Status
        sb.Text = Msg
        txtLog.Text &= vbCrLf & Msg
        txtLog.SelectionStart = txtLog.Text.Length
        txtLog.SelectionLength = 0
        If txtLog.TextLength > 32000 Then txtLog.Text = txtLog.Text.Substring(10000)
        If txtLog.Visible Then txtLog.ScrollToCaret()
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        ' clean up
        m_voice.StopSpeaking()
        m_Word.ExitWord()
        m_voice = Nothing
        m_Word = Nothing
    End Sub

    Private Sub cboEffect_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEffect.DropDown
        If cboEffect.Items.Count = 1 Then
            ' query accessible webservices...
            ShowStatus("Querying WebServices ...")
            cboEffect.DroppedDown = False
            cboEffect.Cursor = Cursors.AppStarting
            cboEffect.Enabled = False
            Try
                ShowStatus("Adding BabelFish effects ...")
                Dim babelfish As New net.xmethods.www.BabelFishService
                Dim ret As String = babelfish.BabelFish("en_fr", "123")
                If TrimWhiteSpace(ret) = "123" Then
                    cboEffect.Items.Add("BabelFish: RoundTrip")
                Else
                   ' The BabelFish webservice was dead for quite some time, but
                   ' appears to be back.  The "123" test is to filter out a
                   ' "this webservice is not active" response that would
                   ' somtimes get returned.
                    ShowStatus(String.Format("Response from BabelFish was '{0}'.", ret))
                End If
            Catch ex As Exception
                ShowStatus("Unable to access BabelFish: " & ex.Message)
            End Try

                '<option value="en_zh" >English to Chinese</option>
                '<option value="en_fr" >English to French</option>
                '<option value="en_de" >English to German</option>
                '<option value="en_it" >English to Italian</option>
                '<option value="en_ja" >English to Japanese</option>
                '<option value="en_ko" >English to Korean</option>
                '<option value="en_pt" >English to Portuguese</option>
                '<option value="en_es" SELECTED>English to Spanish</option>

                '<option value="zh_en" >Chinese to English</option>
                '<option value="fr_en" >French to English</option>
                '<option value="de_en" >German to English</option>
                '<option value="it_en" >Italian to English</option>
                '<option value="ja_en" >Japanese to English</option>
                '<option value="ko_en" >Korean to English</option>
                '<option value="pt_en" >Portuguese to English</option>
                '<option value="ru_en" >Russian to English</option>
                '<option value="es_en" >Spanish to English</option>

                '<option value="de_fr" >German to French</option>
                '<option value="fr_de" >French to German</option>

            ' Too slow!
            Try
                ShowStatus("Adding WebserviceX effects ...")
                Dim webx As New net.webservicex.www.TranslationService
                Dim ret As String = webx.Translate(net.webservicex.www.Language.EnglishTOFrench, "123")
                If ret = "123" Then
                    cboEffect.Items.Add("WebserviceX: RoundTrip")
                End If
            Catch ex As Exception
                ShowStatus("Unable to access WebserviceX: " & ex.Message)
            End Try

            Try
                ShowStatus("Adding Borland effects ...")
                Dim borland As New com.borland.ww6.IBorlandBabelservice
                Dim resp As String = borland.SupportedLanguages
                Dim l As String
                For Each l In resp.Split(vbLf.Chars(0))
                    If l.Trim.Length > 0 Then
                        cboEffect.Items.Add("Borland: " & l)
                    End If
                Next
            Catch ex As Exception
                ShowStatus("Unable to access Borland: " & ex.Message)
            End Try
            Try
                ShowStatus("Adding AspxRunway effects ...")
                Dim aspxrunway As New com.aspxpressway.www.piglatin
                If aspxrunway.toPigLatin("Hello, world") > "" Then
                    cboEffect.Items.Add("AspxRunway: PigLatin")
                End If
            Catch ex As Exception
                ShowStatus("Unable to access AspxRunway: " & ex.Message)
            End Try

            cboEffect.Items.Add("(Random)")
            ShowStatus("Adding Borland effects ...")
            cboEffect.Enabled = True
            cboEffect.Cursor = Cursors.Default
        End If
    End Sub

#Region "ToolWindow Overrides"
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
      Get
         ' This sets the window up as not "stealing" focus (NOACTIVATE)
         ' Unfortunately, the dropdown lists force activation!!
         Const WS_EX_NOACTIVATE As Integer = &H8000000
         Dim Result As CreateParams
         Result = MyBase.CreateParams
         Result.ExStyle = Result.ExStyle Or WS_EX_NOACTIVATE
         Return Result
      End Get
    End Property

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_NCLBUTTONDOWN Then
            If (m.WParam.ToInt32 = HTCAPTION) Then
                ' shuffle window to the top of the stack.   This is because it behaves
                ' somewhat oddly as a toolwindow (it stays in the current z-order, 
                ' so you can actually drag it BEHIND another window!)
                'Me.BringToFront() ' doesn't work!
                Me.TopMost = True
                Application.DoEvents() ' wait for it....
                Me.TopMost = False
            End If
        End If
        MyBase.WndProc(m)  ' allow ancestor to handle it.
    End Sub
#End Region
End Class
