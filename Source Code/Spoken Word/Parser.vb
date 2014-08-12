' More fun effects I haven't gotten around to....
'   tourette's effects ... $h!+!!
'   http://www.tourettesyndrome.net/Files/CommonTics.PDF
'   stuttering effects
'   George Bush effects: words > 10 letters add extra sylillable :)

'research context-free (CF) parsers.  Two good, recent textbooks on the
'subject are Jurafsky & Martin, "Speech & Language Processing," 
'and Allen's "Natural Language Understanding."

Public Module Parser
    ' remove whitespace (space, tab, CR, LF) from (beginning) end of string
    Public Function TrimWhiteSpace(ByVal Text As String, Optional ByVal StartAndEnd As Boolean = False) As String
        Dim WS As Char() = {" "c, vbTab.Chars(0), vbCr.Chars(0), vbLf.Chars(0)}
        If StartAndEnd Then Text = Text.TrimStart(WS)
        Return Text.TrimEnd(WS)
    End Function

    Public Function FixText(ByVal Text As String) As String
        Dim s As String = TrimWhiteSpace(Text, True)
        ' normalize all line breaks to full CRLF
        s = s.Replace(vbCrLf, vbCr)
        s = s.Replace(vbLf, vbCr)
        s = s.Replace(vbCr, vbCrLf)
        Return s
    End Function

    Public Function Decode(ByVal Text As String) As String
        Return System.Web.HttpUtility.HtmlDecode(Text)
    End Function

    Public Function ParseWords(ByVal Text As String, Optional ByVal RetainPunctuation As Boolean = False) As ArrayList
        Dim words As New ArrayList
        If Text.Length = 0 Then Return words
        Dim i As Integer, c As Char
        Dim newword As String = ""
        Dim wasNum, wasUpper, wasLower, wasPunct As Boolean

        For i = 0 To Text.Length - 1
            c = Text.Chars(i)
            If Char.IsLetterOrDigit(c) Then
                If Char.IsDigit(c) Then ' IsNumber can be true for "V" (roman numeral 5)!!
                    If Not wasNum Then
                        ' new word
                        If newword > "" Then words.Add(newword)
                        newword = ""
                        wasNum = True
                        wasUpper = False
                        wasLower = False
                    End If
                Else
                    If wasNum Then
                        ' transition
                        If newword > "" Then words.Add(newword)
                        newword = ""
                    ElseIf wasLower AndAlso Char.IsUpper(c) Then
                        ' word transition: aL
                        If newword > "" Then words.Add(newword)
                        newword = ""
                    ElseIf wasUpper AndAlso Char.IsLower(c) Then
                        ' casing change, might be a word (might not)
                        If newword.Length > 1 Then
                            ' ok ... split it.  steal last letter from oldword first
                            words.Add(newword.Substring(0, newword.Length - 1))
                            newword = newword.Substring(newword.Length - 1)
                        End If
                    ' else same case as before; just accumulate
                    End If
                    wasNum = False
                    wasUpper = Char.IsUpper(c)
                    wasLower = Char.IsLower(c)
                End If
                newword &= c
                wasPunct = False
            Else ' any whitespace, punctuation or other symbol
                If RetainPunctuation AndAlso Char.IsPunctuation(c) Then
                    ' accumulate as its own "word"
                    If Not wasPunct Then
                        If newword > "" Then words.Add(newword)
                        newword = ""
                    End If
                    wasPunct = True
                Else
                    ' whitespace, control chars, etc. don't accumulate
                    If newword > "" Then words.Add(newword)
                    newword = ""
                    wasPunct = False
                End If
                wasNum = False
                wasUpper = False
                wasLower = False
            End If
        Next

        If newword > "" Then words.Add(newword)
        Return words
    End Function

    Public Function PreProcess(ByVal Text As String) As String
        ' bold, italics, strikethru, underline should be parsed from RTF...
        ' _ = " ", break words
        ' <, >, <=, >=, <>, != : ; ...

        ' Would be better to externalize (xml) these, just never got around to it.
        ' Then the end user could add new entries to the "dictionary".
        Text = Text.Replace(vbCrLf, "." & vbCrLf)
        Text = Text.Replace("   ", " -- ")

        Text = Text.Replace(Chr(147), """") ' 147 = “
        Text = Text.Replace(Chr(148), """") ' 148 = ”
        Text = Text.Replace(Chr(133), "...") ' 148 = …
        Text = Text.Replace(Chr(148), """") ' 148 = ”
        Text = Text.Replace(Chr(146), "'") ' 148 = ’
        Text = Text.Replace("—", " -- ") ' em dash
        Text = Text.Replace("–", "--") ' no, not the same char: 150 -> 45
        Text = Text.Replace(" © ", " copyright ") ' pronounce if separate
        Text = Text.Replace("©", " ") ' ignore if part of name
        Text = Text.Replace(" ® ", " registered ") ' pronounce if separate
        Text = Text.Replace("®", " ") ' ignore if part of name
        Text = Text.Replace(" ™ ", " trademark ") ' pronounce if separate
        Text = Text.Replace("™", " ") ' ignore if part of name
        Text = Text.Replace("§", ".") ' sometimes come from bullet lists

        Text = Text.Replace("_", " ")
        Text = Text.Replace("~", " approximately ")

        ' can't have these because of SAPI XML:
        Text = Text.Replace("->", " arrow ")
        Text = Text.Replace("<>", " not equal to ")
        Text = Text.Replace("<=", " less than or equal to ")
        Text = Text.Replace(">=", " greater than or equal to ")
        Text = Text.Replace("!=", " not equal ")
        Text = Text.Replace("<", " less than ")
        Text = Text.Replace(">", " greater than ")

        Text = Text.Replace(" to,", " too,") ' has trouble with trailing "to's".

        ' collapse multiple tabs and spaces
        Dim len As Integer
        Do
            len = Text.Length
            Text = Text.Replace("  ", " ")
            Text = Text.Replace(vbTab & vbTab, vbTab)
        Loop Until Text.Length = len ' no more substitutions
        Text = Text.Replace(vbTab, " ... ") ' this will become a pause

        Return Text
    End Function

    Public Function PostProcess(ByVal Text As String) As String
        ' Here is where some crazy RegEx processing could occur:
        '    acronyms -> S. Q. L. (more understandable)
        '    long all caps (e.g. YELLING) is just emphatic (<emph>yelling</emph>)
        '    word(s), process(es) should pronounce as "word-Z", "process-EZ"
        '    however, "f(x)" typically read "f of x"

        ' Also, for developer tech docs, expand CamelCasing to separate words ("Camel Casing").
        ' This is supported by the ParseWords fxn, but not currently implemented.

        ' nothing currently implemented...
        Return Text
    End Function
End Module
