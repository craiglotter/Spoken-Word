Imports Word
Imports System.Windows.Forms
Imports System.Threading

Public Class WordHelper
    Private app As Word.Application
    Private currDoc As Word.Document
    Private initialStart, initialEnd As Integer

    ' This is the "increment" we use to navigate a word doc:
    Private Const SELECT_UNIT As Word.WdUnits = WdUnits.wdSentence

    Public Event Status(ByVal Msg As String)

    <System.STAThread()> Private Sub InitWord()
        ExitWord() ' clean up old ref, if any.
        System.Windows.Forms.Application.DoEvents()
        Try
            ' check and see if instance already open.
            app = CType(GetObject(Nothing, "Word.Application"), Word.Application)
            System.Windows.Forms.Application.DoEvents()
        Catch ex As Exception
            RaiseEvent Status("Error in InitWord: " & ex.Message)
        End Try
    End Sub

    <System.STAThread()> Public Sub ExitWord()
        If app Is Nothing Then Exit Sub
        Try
            ' only close word if currently hidden!
            If app.Visible Then Exit Sub
            app.Quit()
            ' decrement COM reference
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app)
        Catch ex As Exception
            RaiseEvent Status("Error in ExitWord: " & ex.Message)
        End Try
        app = Nothing
    End Sub

    Public ReadOnly Property CurrentDocument() As Word.Document
        Get
            InitWord()
            If app Is Nothing Then Return Nothing
            If Not currDoc Is Nothing Then
                Try
                    ' test to see if prior doc is still available....
                    Dim foo As String = currDoc.Name
                Catch ex As Exception
                    ' no doc, or doc has been closed.
                    RaiseEvent Status("Error in CurrentDocument: " & ex.Message)
                    currDoc = Nothing
                End Try
            End If
            If currDoc Is Nothing Then
                System.Windows.Forms.Application.DoEvents()
                Try
                    currDoc = app.ActiveDocument
                Catch ex As Exception
                    RaiseEvent Status("Error in CurrentDocument: " & ex.Message)
                    Return Nothing
                End Try
            End If

            Return currDoc
        End Get
    End Property

    Public Sub SetLimits()
        initialStart = 0
        initialEnd = Integer.MaxValue
        Dim sel As Word.Selection = GetSelection()
        If Not sel Is Nothing Then
            initialStart = sel.Range.Start
            initialEnd = sel.Range.End
            If initialStart = initialEnd Then
                'initialEnd = Integer.MaxValue
                ' find the end of the story (max selection range)
                Dim story As Word.Range = currDoc.Range
                story.WholeStory()
                initialEnd = story.End
            End If
        End If
    End Sub

    Public Function GetSelection() As Word.Selection
        Dim doc As Word.Document = CurrentDocument
        If doc Is Nothing Then Return Nothing
        Try
            Return doc.ActiveWindow.Selection
        Catch
            Return Nothing
        End Try
    End Function

    Public Function IsAnIndexBlock(ByVal sel As Word.Selection) As Boolean
        Return False

        ' The intent of this function was to detect if the selection is 
        ' a Table of Contents, Table of Figures, Index, Glossary, etc.
        ' and therefore something that does not need to be "read".
        ' Unfortunatley, I don't know how to figgure that out,
        ' and lost interest after I implemented "Skip".

        'Dim sy As Word.Style = sel.ParagraphFormat.Style
        'Console.WriteLine(sy.NameLocal & " " & sy.BuiltIn & " " & sy.ListLevelNumber)
        'Dim s As Word.WdBuiltinStyle
        'Select Case s
        '    Case Word.WdBuiltinStyle.wdStyleEmphasis
        '    Case Word.WdBuiltinStyle.wdStyleHeading9 To Word.WdBuiltinStyle.wdStyleHeading1
        '    Case Word.WdBuiltinStyle.wdStyleIndex9 To Word.WdBuiltinStyle.wdStyleIndex1
        '    Case Word.WdBuiltinStyle.wdStyleListNumber
        'End Select
    End Function

    ' Not currently used, but the intent was to pre-process the next block,
    ' rather than waiting until it has already been read.  Would cut down on
    ' lags between sentences when using effects, but those are really just novelties!
    Public Function PeekNext() As Word.Range
        Dim sel As Word.Selection = GetSelection()
        If sel Is Nothing Then Return Nothing
        Dim r As Word.Range = sel.Range
        r.Start = r.End
        Try
            r.Collapse()
        Catch ex As Exception
            ' occasionally throws an exception "call was rejected by the Callee"?!?!
            RaiseEvent Status("Error in PeekNext: " & ex.Message)
        End Try
        r.End += 1
        If r.End >= initialEnd Then Return Nothing
        r.Expand(SELECT_UNIT)
        If r.End > r.Start Then Return r
        Return Nothing
    End Function

    ' set caret at specified location
    Public Sub SetCaret(ByVal Pos As Integer)
        Dim sel As Word.Selection = GetSelection()
        If sel Is Nothing Then Exit Sub
        If sel.Start = Pos Then Exit Sub
        RaiseEvent Status("SetCaret: Desired position is " & Pos & ", Actual is " & sel.Start)
        sel.Start = Pos
        sel.End = Pos
        Try
            sel.Collapse()
        Catch ' useless errors.
        End Try
    End Sub

    ' sets caret AND selects next block
    Public Sub MoveTo(ByVal Pos As Integer)
        SetCaret(Pos)
        NavNext()
    End Sub

    ' navigate in Word to next block of text.
    Public Function NavNext(Optional ByVal Recurse As Boolean = True) As Boolean
        Dim sel As Word.Selection = GetSelection()
        If sel Is Nothing Then Return False
        Dim oldstart As Integer = sel.Range.Start
        Dim oldEnd As Integer = sel.Range.End
        Dim i As Integer = 1
        Do ' don't ask.  sometimes it gets stuck.
            sel.Start = sel.End
            Try
                sel.Collapse()
            Catch ' useless errors.
            End Try
            sel.End += i
            If sel.End >= initialEnd Then
                ' at end of original selection.
                sel.End = initialEnd
                Exit Do
            End If
            i += 1
            sel.Expand(SELECT_UNIT)
        Loop Until sel.End > oldEnd                               ' YES, ALL of these
        If sel.Start < initialStart Then sel.Start = initialStart ' are necessary.
        If sel.Start < oldstart Then sel.Start = oldstart '       ' Navigation in Word 
        If sel.Start < oldEnd Then sel.Start = oldEnd '           ' can be a pain-in-the
        If sel.Start < oldEnd Then '                              ' (you-know-what)!
            ' This happens sometimes in a table cell ... it refuses to select a portion.
            ' have to decrement the end as well to make it work.
            sel.End -= 1
            sel.Start = oldEnd
            If sel.Start = sel.End Then
                sel.Expand(SELECT_UNIT) ' ok ... now we're stuck in a cell.  expand to get it all THEN navnext.
                If Recurse Then Return NavNext(False) Else Return False
            End If
        End If

        ' Force doc to show selection.
        sel.Document.ActiveWindow.ScrollIntoView(sel.Range)
        If sel.End > sel.Start Then Return True
        Return False
    End Function

    ' navigate in Word to prior block of text.  Note it is much simpler than NavNext!
    Public Function NavPrior() As Boolean
        Dim sel As Word.Selection = GetSelection()
        If sel Is Nothing Then Return False
        If sel.Start = 0 Then Return False
        Dim currstart As Integer = sel.Start
        sel.Start -= 2
        sel.End = sel.Start + 1
        sel.Expand(SELECT_UNIT)
        If sel.Start <= initialStart Then Return False
        If sel.Start = currstart Then Return False
        sel.Document.ActiveWindow.ScrollIntoView(sel.Range)
        If sel.Range.End > sel.Range.Start Then Return True
        Return False
    End Function
End Class
