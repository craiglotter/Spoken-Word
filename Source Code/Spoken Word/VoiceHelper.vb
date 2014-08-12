Imports SpeechLib

' This class wraps the SpeechLib classes (for reusability)

Public Class VoiceHelper
    Private WithEvents m_voice As SpeechLib.SpVoiceClass
    Public Event Bookmark(ByVal Bookmark As String, ByVal BookmarkId As Integer)
    Public Event Status(ByVal Msg As String)

    Private m_SuppressBookmarks As Boolean = False
    Private m_IsSpeaking As Boolean = False
    Private m_IsPaused As Boolean = False
    Private m_Speed As Integer
    Private m_Pitch As Integer
    Private m_volume As Integer
    Private m_Speaker As String
    Private m_LastPhrase As String = ""
    Private m_LastBookmark As Integer = 0

    Public Sub New()
        m_voice = New SpeechLib.SpVoiceClass
        ' get defaults
        m_Speed = m_voice.Rate
        m_volume = m_voice.Volume
        m_Pitch = 0 '?? I think this is the default for all
        m_Speaker = CurrentVoice
    End Sub

    ' returns installed voices (just their names)
    Public ReadOnly Property Voices() As ArrayList
        Get
            Dim v As New ArrayList
            Dim token As SpeechLib.ISpeechObjectToken
            For Each token In m_voice.GetVoices()
                v.Add(token.GetDescription)
            Next
            Return v
        End Get
    End Property

    Public ReadOnly Property CurrentVoice() As String
        Get
            Return m_voice.Voice.GetDescription
        End Get
    End Property

    Public Sub SetVoice(ByVal Name As String)
        m_Speaker = Name
    End Sub

    Public Function SAPIfy(ByVal Text As String) As String
        ' reformat selection for SAPI...

        ' This is where I wanted to do stuff based on the format 
        ' of the text in Word -- "announce" headers and titles 
        ' (maybe in a different voice, or at least a different pitch);
        ' add emphasis tags for words in bold or italics or underlined....
        ' Like other parts, I haven't gotten to it yet.....

        ' ... so all it does currently is add "moments of silence" where needed.
        Text = Text.Replace("--", "<silence msec=""250""/> , ")
        Text = Text.Replace("...", "<silence msec=""400""/>")
        Return Text
    End Function

    Public Sub Speak(ByVal Text As String, ByVal Bookmark As Integer, Optional ByVal ResetParams As Boolean = True)  
        If ResetParams Then ' need to wait until IsSpeaking = False
            If m_voice.Rate <> m_Speed OrElse m_voice.Volume <> m_volume OrElse m_voice.Voice.GetDescription <> m_Speaker Then
                ' only have to wait if we ACTUALLY need to reset parameters.
                ' NOTE: this can probably cause strange bugs, since the voice has a queue
                ' and the app is also "queuing" and we might get a repeat or skip or back
                ' request *while* this is waiting ... resulting in out-of-sequence playback.
                RaiseEvent Status("Waiting to update voice parameters...")
                m_SuppressBookmarks = True
                Dim i As Integer = 0
                Do Until IsSpeaking = False
                    i += 1
                    Application.DoEvents()
                    If i > 10000 Then Exit Do
                Loop
                m_SuppressBookmarks = False
                ' Set new voice
                If m_voice.Voice.GetDescription <> m_Speaker Then
                    Dim token As SpeechLib.ISpeechObjectToken
                    For Each token In m_voice.GetVoices()
                        If token.GetDescription = m_Speaker Then
                            m_voice.SetVoice(CType(token, SpeechLib.ISpObjectToken))
                        End If
                    Next
                End If
                m_voice.Rate = m_Speed
                m_voice.Volume = m_volume
            End If
        End If
        RaiseEvent Status("Queuing new text to speak.")
        m_LastPhrase = Text
        m_LastBookmark = Bookmark
        Text = "<SAPI><PITCH MIDDLE=""" & m_Pitch & """><BOOKMARK MARK=""-1""/>" & Text & "<BOOKMARK MARK=""" & Bookmark & """/></PITCH></SAPI>"
        'm_IsSpeaking = True
        m_voice.Speak(Text, SpeechVoiceSpeakFlags.SVSFDefault Or SpeechVoiceSpeakFlags.SVSFIsXML Or SpeechVoiceSpeakFlags.SVSFlagsAsync)
    End Sub

    ' Huh?  What'd he say?
    Public Sub Repeat()
        Dim r As Integer = m_voice.Rate - 2
        If r < -6 Then r = -6
        Dim v As Integer = CInt(m_voice.Volume + 5)
        v = Math.Min(v, 100)
        StopSpeaking()
        m_voice.Rate = r ' slow down!
        m_voice.Volume = v ' speak up!
        Speak(m_LastPhrase, m_LastBookmark, False)
    End Sub

    Public ReadOnly Property IsSpeaking() As Boolean
        Get
            Return m_IsSpeaking
        End Get
    End Property

    Public Sub StopSpeaking()
        ' Voice interface has NO means for clearing the speech queue,
        ' so we have to SKIP through current playback to clear it.
        m_SuppressBookmarks = True
        Dim i As Integer = 0
        RaiseEvent Status("Attempting to stop speaking...")
        If m_IsPaused Then ResumeSpeaking()
        Do Until IsSpeaking = False
            i += 1
            RaiseEvent Status("skipping...")
            m_voice.Skip("SENTENCE", 5)
            Application.DoEvents()
            If i > 100 Then Exit Do ' >500 sentences queued up!?!  Bail out!
        Loop
        m_SuppressBookmarks = False
        RaiseEvent Status("Speaking stopped.")
    End Sub

    Public Sub PauseSpeaking()
        If m_IsPaused Then Exit Sub
        RaiseEvent Status("Pausing speaking...")
        m_voice.Pause()
        m_IsPaused = True
    End Sub

    Public Sub ResumeSpeaking()
        If Not m_IsPaused Then Exit Sub
        RaiseEvent Status("Resuming speaking...")
        m_voice.Resume()
        m_IsPaused = False
    End Sub

    ' Pitch has range of -10 to +10
    Public Property Pitch() As Integer
        Get
            Return m_Pitch
        End Get
        Set(ByVal Value As Integer)
            Value = Math.Min(Value, 10)
            Value = Math.Max(Value, -10)
            m_Pitch = Value
        End Set
    End Property

    ' Speed has range of -10 to +10
    Public Property Speed() As Integer
        Get
            Return m_Speed
        End Get
        Set(ByVal Value As Integer)
            Value = Math.Min(Value, 10)
            Value = Math.Max(Value, -10)
            m_Speed = Value
        End Set
    End Property

    ' Volume has range of 0 to 100
    Public Property Volume() As Integer
        Get
            Return m_volume
        End Get
        Set(ByVal Value As Integer)
            Value = Math.Min(Value, 100)
            Value = Math.Max(Value, 0)
            m_volume = Value
        End Set
    End Property

    Private Sub m_voice_Bookmark(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal Bookmark As String, ByVal BookmarkId As Integer) Handles m_voice.Bookmark
        If m_SuppressBookmarks Then Exit Sub
        RaiseEvent Status("Bookmark encountered: " & Bookmark)
        RaiseEvent Bookmark(Bookmark, BookmarkId)
    End Sub

    Private Sub m_voice_StartStream(ByVal StreamNumber As Integer, ByVal StreamPosition As Object) Handles m_voice.StartStream
        m_IsSpeaking = True
        RaiseEvent Status("Start of stream.")
    End Sub

    Private Sub m_voice_EndStream(ByVal StreamNumber As Integer, ByVal StreamPosition As Object) Handles m_voice.EndStream
        m_IsSpeaking = False
        RaiseEvent Status("End of stream.")
    End Sub
End Class
