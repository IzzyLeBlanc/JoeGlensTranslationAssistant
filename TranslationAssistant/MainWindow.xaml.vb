
Class MainWindow
    Public rawTextArray() As String
    Dim translatedTextArray() As String
    Dim topText As String
    Dim bottomText As String
    Public arrayPointer As Integer
    Dim parseSentence() As String
    Public glossaryArray(,) As String
    Dim replaced As Boolean
    Dim parsenumber, parsepointer As Integer
    Dim parsers() As String
    Dim filepath As String
    Dim txtOutput As String
    Dim totalRawLines As Integer
    Dim tlComplete As Integer
    Dim caretIndex As Integer = 0
    Dim linenumber(,) As Integer
    Dim synthesizer As Object

    Public tmrCopyToClipboard As System.Windows.Threading.DispatcherTimer = New System.Windows.Threading.DispatcherTimer
    Public tmrFileSaved As System.Windows.Threading.DispatcherTimer = New System.Windows.Threading.DispatcherTimer
    Public tmrAutoSave As System.Windows.Threading.DispatcherTimer = New System.Windows.Threading.DispatcherTimer

    Public Const CF_TEXT = 1   ' ANSI text
    Public Const CF_OEMTEXT = 7
    Public Const CF_UNICODETEXT = 13
    Public Const GMEM_SHARE = &H2000&
    Public Const GMEM_MOVEABLE = &H2
    Public Const GMEM_ZEROINIT = &H40
    Public Const FOR_CLIPBOARD = GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT

    Private Sub TextContent(ByVal value As String)
        Dim splitRawTL(1) As String

        splitRawTL = Split(value, "---SEPERATOR---", 2)
        Erase rawTextArray
        Erase translatedTextArray
        txtOutput = splitRawTL(0) & "---SEPERATOR---" & vbNewLine
        rawTextArray = Split(splitRawTL(0), vbNewLine)
        translatedTextArray = Split(Mid(splitRawTL(1), 3, splitRawTL(1).Length - 2), vbNewLine)
        While (1)
            If rawTextArray(rawTextArray.GetUpperBound(0)) = "" OrElse rawTextArray(rawTextArray.GetUpperBound(0)) = Nothing Then
                ReDim Preserve rawTextArray(rawTextArray.GetUpperBound(0) - 1)
                ReDim Preserve translatedTextArray(rawTextArray.GetUpperBound(0))
            Else
                Exit While
            End If
        End While
        For i As Integer = 0 To rawTextArray.Count - 1
            rawTextArray(i) = Replace(rawTextArray(i), vbNewLine, "")
            rawTextArray(i) = Replace(rawTextArray(i), vbLf, "")
        Next
        reviewTop.Text = ""
        ReDim linenumber(1, rawTextArray.GetUpperBound(0))
        currentRawLine.Text = replaceAndParse(rawTextArray(0))
        currentTranslatedLine.Text = translatedTextArray(0)
        reviewBottom.Text = updateReview(1, rawTextArray.GetUpperBound(0))
        arrayPointer = 0
        LineStatus.Content = "Line: 1/" & rawTextArray.Length.ToString
        Dim rawTotal As New List(Of String)
        rawTotal.AddRange(rawTextArray)
        rawTotal.RemoveAll(Function(str) String.IsNullOrEmpty(str))
        totalRawLines = rawTotal.Count
        calculateProgress()
        currentRawLine.Focus()
        currentTranslatedLine.Focus()
        Task.WaitAll()
        writeToClipboard(currentRawLine.Text)
        menuSave.IsEnabled = True
        menuClipboard.IsEnabled = True
        text2Speech(currentRawLine.Text)
    End Sub

    Public Sub glossaryContent(ByVal value As List(Of String))
        Try
            ReDim glossaryArray(0, 0)
            If value.Count > 0 Then
                value.RemoveAll(Function(str) String.IsNullOrEmpty(str))
                ReDim glossaryArray(value.Count - 1, 1)
                For i As Integer = 0 To value.Count - 1
                    Dim splitarray() As String = Split(value(i), ",", 2)
                    glossaryArray(i, 0) = splitarray(0)
                    If (splitarray.GetUpperBound(0)) = 1 Then
                        glossaryArray(i, 1) = splitarray(1)
                    Else
                        glossaryArray(i, 1) = ""
                    End If
                Next
                If rawTextArray IsNot Nothing Then
                    currentRawLine.Text = replaceAndParse(rawTextArray(arrayPointer))
                    writeToClipboard(currentRawLine.Text)
                End If
            End If
        Catch eX As Exception
            MsgBox("Something is wrong with the Profile, error Code: " & eX.ToString, vbCritical, "Unexpected Error")
        End Try
    End Sub

    Private Sub writeToClipboard(ByVal inputText As String)
        Dim ptr As IntPtr = System.Runtime.InteropServices.Marshal.StringToHGlobalAuto(inputText)
        While (1)
            If NativeMethods.OpenClipboard(Process.GetCurrentProcess.Handle) = True Then
                NativeMethods.EmptyClipboard()
                NativeMethods.SetClipboardData(CF_UNICODETEXT, ptr)
                NativeMethods.CloseClipboard()
                Exit While
            End If
        End While
    End Sub

    Public Function replaceAndParse(ByVal textInput As String) As String
        Dim txtbuffer As String = Replace(textInput, "$", "")
        txtbuffer = Replace(txtbuffer, "%", "")
        replaced = False
        If glossaryArray IsNot Nothing Then
            For i As Integer = 0 To glossaryArray.GetUpperBound(0)
                If glossaryArray(i, 0) IsNot Nothing AndAlso txtbuffer Like "*" & glossaryArray(i, 0) & "*" Then
                    txtbuffer = txtbuffer.Replace(glossaryArray(i, 0), glossaryArray(i, 1))
                    replaced = True
                End If
            Next
            parsenumber = parseCount(txtbuffer)
            parsepointer = -1
            Return txtbuffer
        Else
            parsenumber = parseCount(txtbuffer)
            parsepointer = -1
            Return txtbuffer
        End If
    End Function

    Private Sub calculateProgress()
        Dim count As Integer = 0
        Dim wordCount As Integer = 0

        For i As Integer = 0 To rawTextArray.GetUpperBound(0)
            If rawTextArray(i) <> "" AndAlso translatedTextArray(i) <> "" Then
                count = count + 1
            End If
            wordCount = wordCount + (Split(translatedTextArray(i), " ").Count)
        Next
        tlComplete = (count / totalRawLines) * 100
        completionStatus.Content = FormatPercent(count / totalRawLines, 0) & " Complete"
        wordCountStatus.Content = wordCount.ToString & " Words"
    End Sub

    Private Function parseCount(ByVal inputText As String) As Integer
        Erase parseSentence
        parseSentence = inputText.Split(parsers, StringSplitOptions.RemoveEmptyEntries)
        Return parseSentence.Length
    End Function

    Private Function updateReview(ByVal startIndex As Integer, ByVal endIndex As Integer) As String
        Dim count = startIndex
        Dim endcount = endIndex
        Dim displayText As String = ""

        While (count <= endcount)
            If rawTextArray(count) <> "" Then
                Dim i As Integer = 0
                While (1)
                    linenumber(0, count + i) = displayText.Length
                    rawTextArray(count + i) = Replace(rawTextArray(count + i), "%", "")
                    displayText = displayText & Replace(rawTextArray(count + i), "$", "")
                    linenumber(1, count + i) = displayText.Length
                    i = i + 1
                    If count + i > rawTextArray.GetUpperBound(0) OrElse count + i > endcount OrElse Mid(rawTextArray(count + i), 1, 1) <> "$" Then
                        displayText = displayText & vbNewLine
                        For x As Integer = 1 To i
                            displayText = displayText & translatedTextArray(count + (x - 1)) & Space(1)
                        Next
                        displayText = displayText & vbNewLine & vbNewLine
                        count = count + i
                        Exit While
                    End If
                End While
            Else
                displayText = displayText & vbNewLine
                count = count + 1
            End If
        End While
        Return displayText
    End Function

    Private Function convertToList(ByVal value As String) As List(Of String)
        Dim result As List(Of String)

        result = value.Split(vbNewLine).ToList
        For i As Integer = 0 To result.Count - 1
            result(i) = Replace(result(i), vbLf, "")
        Next
        result.RemoveAll(Function(str) String.IsNullOrEmpty(str))
        Return result
    End Function

    Private Sub insertPunctuation(ByVal index As Integer)
        Dim Response() As String = {"「」", "『』", "【】", "…", "〜", "〈〉", "《》", "ー"}
        Dim cursorStart As Integer = currentTranslatedLine.CaretIndex
        currentTranslatedLine.Text = currentTranslatedLine.Text.Insert(cursorStart, Response(index))
        currentTranslatedLine.Focus()
        currentTranslatedLine.CaretIndex = cursorStart + 1
    End Sub

    Private Sub AddToDictionary(ByVal sender As Object, ByVal e As EventArgs)
        Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex", True)
        fileopened.WriteLine(currentTranslatedLine.SelectedText)
        fileopened.Close()
        MsgBox("The word """ & currentTranslatedLine.SelectedText & """ has been added to the dictionary", vbOKOnly, "Dictionary")
    End Sub

    Private Sub tmrTick()
        tmrCopyToClipboard.IsEnabled = False
        If currentRawLine.SelectionLength = 0 Then
            writeToClipboard(currentRawLine.Text)
        Else
            writeToClipboard(currentRawLine.SelectedText)
        End If
    End Sub

    Private Sub tmrFileSave()
        tmrFileSaved.IsEnabled = False
        filesaved.Content = ""
    End Sub
    Private Sub ConfigureTTS()
        If menuTTSJP.IsChecked = True Then
            synthesizer.SelectVoice("Microsoft Server Speech Text to Speech Voice (ja-JP, Haruka)")
        ElseIf menuTTSCN.IsChecked = True Then
            synthesizer.SelectVoice("Microsoft Server Speech Text to Speech Voice (zh-CN, HuiHui)")
        End If
        synthesizer.Volume = 100
        synthesizer.Rate = -2
        synthesizer.SetOutputToDefaultAudioDevice()
    End Sub

    Private Sub text2Speech(ByVal text As String)
        If My.Settings.TTS = True Then
            If synthesizer.State = Microsoft.Speech.Synthesis.SynthesizerState.Speaking Then
                synthesizer.SpeakAsyncCancelAll()
            End If
            synthesizer.SpeakAsync(text)
        End If
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim dictionary As IList = SpellCheck.GetCustomDictionaries(currentTranslatedLine)
        parsers = Split(My.Settings.ParseChar, " ")
        filesaved.Content = ""
        tmrCopyToClipboard.IsEnabled = False
        tmrCopyToClipboard.Interval = TimeSpan.FromMilliseconds(400)
        tmrFileSaved.IsEnabled = False
        tmrFileSaved.Interval = TimeSpan.FromMilliseconds(2000)
        tmrAutoSave.IsEnabled = False
        AddHandler tmrCopyToClipboard.Tick, AddressOf tmrTick
        AddHandler tmrFileSaved.Tick, AddressOf tmrFileSave
        If My.Computer.FileSystem.DirectoryExists(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile") Then
            If My.Computer.FileSystem.FileExists(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".csv") Then
                Dim fileOpened As New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".csv")
                glossaryContent(convertToList(fileOpened.ReadToEnd))
                fileOpened.Close()
            Else
                MsgBox("Profile named " & My.Settings.ProfileUsed & " Is missing. Using default profile", vbCritical, "File Missing")
                If My.Computer.FileSystem.FileExists(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\Default.csv") Then
                    Dim fileOpened As New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\Default.csv")
                    glossaryContent(convertToList(fileOpened.ReadToEnd))
                    fileOpened.Close()
                Else
                    MsgBox("Default profile Is missing Or corrupted. Creating New blank Default Profile", vbCritical, "File Missing")
                    Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\Default.csv", False)
                    fileopened.WriteLine("")
                    fileopened.Close()
                End If
                My.Settings.ProfileUsed = "Default"
            End If
            If My.Computer.FileSystem.FileExists(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex") Then
                dictionary.Add(New Uri(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex"))
            Else
                Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex", False)
                fileopened.WriteLine("#LID 1033")
                fileopened.Close()
                dictionary.Add(New Uri(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex"))
            End If
        Else
            My.Computer.FileSystem.CreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile")
            Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\Default.csv", False)
            fileopened.WriteLine("")
            fileopened.Close()
            fileopened = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\Default.lex", False)
            fileopened.WriteLine("#LID 1033")
            fileopened.Close()
            dictionary.Add(New Uri(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex"))

        End If

        menuProgress.IsChecked = My.Settings.ShowProgress
        menuOnTop.IsChecked = My.Settings.OnTop
        If menuProgress.IsChecked = False Then
            completionStatus.Visibility = True
            LineStatus.Visibility = True
        End If
        If menuOnTop.IsChecked = True Then
            Me.Topmost = True
        End If

        Try
            synthesizer = New Microsoft.Speech.Synthesis.SpeechSynthesizer
            For Each voice As Microsoft.Speech.Synthesis.InstalledVoice In synthesizer.GetInstalledVoices
                If voice.VoiceInfo.Name = "Microsoft Server Speech Text to Speech Voice (ja-JP, Haruka)" Then
                    menuTTSJP.IsEnabled = True
                ElseIf voice.VoiceInfo.Name = "Microsoft Server Speech Text to Speech Voice (zh-CN, HuiHui)" Then
                    menuTTSCN.IsEnabled = True
                End If
            Next
            If My.Settings.TTS = True Then
                If My.Settings.TTSLang = 0 And menuTTSJP.IsEnabled = True Then
                    menuTTSJP.IsChecked = True
                    ConfigureTTS()
                ElseIf My.Settings.TTSLang = 1 And menuTTSJP.IsEnabled = True Then
                    menuTTSCN.IsChecked = True
                    ConfigureTTS()
                Else
                    My.Settings.TTS = False
                End If
            End If
        Catch
            menuTTS.Header = "No TTS Engine Installed"
            menuTTS.IsEnabled = False
            My.Settings.TTS = False
        End Try
        Task.WaitAll()
    End Sub

    Private Sub MainWindow_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        reviewTop.Height = 320 + ((Me.ActualHeight - 686) / 2)
        If Me.ActualHeight > 681 Then
            reviewBottom.Height = 130 + ((Me.ActualHeight - 686) / 2)
            reviewTop.Height = 320 + ((Me.ActualHeight - 686) / 2)
        Else
            reviewTop.Height = 320 + ((Me.ActualHeight - 686) / 1)
            reviewBottom.Height = 130
        End If
    End Sub

    Private Sub menuOpen_Click(sender As Object, e As EventArgs) Handles menuOpen.Click
        Dim inputText As String
        Me.Topmost = False
        Dim ofdOpen As New Microsoft.Win32.OpenFileDialog()
        ofdOpen.Filter = "Text Files (*.txt)|*.txt"

        If ofdOpen.ShowDialog = True AndAlso ofdOpen.FileName <> "" Then
            filepath = ofdOpen.FileName
            Dim fileopened As New System.IO.StreamReader(filepath)
            inputText = fileopened.ReadToEnd
            fileopened.Close()
            If inputText.Contains("---SEPERATOR---") Then
                TextContent(inputText)
            Else
                MsgBox("The file you have chosen Is Not supported by this app", vbCritical, "File Error")
            End If
        End If
        If menuOnTop.IsChecked Then
            Me.Topmost = True
        End If
    End Sub

    Private Sub menuSave_Click(sender As Object, e As EventArgs) Handles menuSave.Click
        translatedTextArray(arrayPointer) = currentTranslatedLine.Text
        Dim fileout As String = txtOutput
        For i As Integer = 0 To translatedTextArray.GetUpperBound(0)
            fileout = fileout & translatedTextArray(i) & vbNewLine
        Next
        Dim fileoutput As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(filepath, False)
        Try
            fileoutput.Write(fileout)
            fileoutput.Close()
            filesaved.Content = "File saved...."
            tmrFileSaved.IsEnabled = True
        Catch ex As Exception
            MsgBox("File saving error. Check cuurent file Is Not used by other apps", vbCritical, "File Write Error")
        End Try
    End Sub

    Private Sub menuNew_Click(sender As Object, e As RoutedEventArgs) Handles menuNew.Click
        Dim newWindow As New frmNew
        Me.Topmost = False
        newWindow.ShowDialog()
        If newWindow.DialogResult = True Then
            TextContent(newWindow.rawOutputText)
            filepath = newWindow.filename
        End If
        If menuOnTop.IsChecked Then
            Me.Topmost = True
        End If
    End Sub

    Private Sub menuPhrase_Click(sender As Object, e As RoutedEventArgs) Handles menuPhrase.Click
        Dim newWindow As New frmPhrase
        Me.Topmost = False
        If newWindow.ShowDialog = True Then
            glossaryContent(convertToList(newWindow.glossaryText))
        End If
        If menuOnTop.IsChecked Then
            Me.Topmost = True
        End If
    End Sub

    Private Sub menuProfile_Click(sender As Object, e As RoutedEventArgs) Handles menuProfile.Click
        Dim newWindow As New frmProfile
        Me.Topmost = False
        If newWindow.ShowDialog = True Then
            glossaryContent(newWindow.textOutput)
            parsers = Split(My.Settings.ParseChar, " ")
            Dim dictionary As IList = SpellCheck.GetCustomDictionaries(currentTranslatedLine)
            dictionary.Clear()
            If My.Computer.FileSystem.FileExists(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex") = False Then
                Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex", False)
                fileopened.WriteLine("#LID 1033")
                fileopened.Close()
            End If
            dictionary.Add(New Uri(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex"))
        End If
        If menuOnTop.IsChecked Then
            Me.Topmost = True
        End If
    End Sub

    Private Sub menuProgress_Click(sender As Object, e As RoutedEventArgs) Handles menuProgress.Click
        If menuProgress.IsChecked Then
            menuProgress.IsChecked = False
            completionStatus.Visibility = True
            LineStatus.Visibility = True
        Else
            menuProgress.IsChecked = True
            completionStatus.Visibility = False
            LineStatus.Visibility = False
        End If
        My.Settings.ShowProgress = menuProgress.IsChecked
        My.Settings.Save()
    End Sub

    Private Sub menuOnTop_Click(sender As Object, e As RoutedEventArgs) Handles menuOnTop.Click
        If menuOnTop.IsChecked = True Then
            menuOnTop.IsChecked = False
            Me.Topmost = False
        Else
            menuOnTop.IsChecked = True
            Me.Topmost = True
        End If

        My.Settings.OnTop = menuOnTop.IsChecked
        My.Settings.Save()
    End Sub
    Private Sub menuTTSJP_Click(sender As Object, e As RoutedEventArgs) Handles menuTTSJP.Click
        If menuTTSJP.IsChecked = True Then
            menuTTSJP.IsChecked = False
            My.Settings.TTS = False
        Else
            menuTTSJP.IsChecked = True
            menuTTSCN.IsChecked = False
            My.Settings.TTS = True
            My.Settings.TTSLang = 0
            ConfigureTTS()
        End If

        My.Settings.Save()
    End Sub
    Private Sub menuTTSCN_Click(sender As Object, e As RoutedEventArgs) Handles menuTTSCN.Click
        If menuTTSCN.IsChecked = True Then
            menuTTSCN.IsChecked = False
            My.Settings.TTS = False
        Else
            menuTTSCN.IsChecked = True
            menuTTSJP.IsChecked = False
            My.Settings.TTS = True
            My.Settings.TTSLang = 1
            ConfigureTTS()
        End If

        My.Settings.Save()
    End Sub

    Private Sub menuSingle_Click(sender As Object, e As RoutedEventArgs) Handles menuSingle.Click
        insertPunctuation(0)
    End Sub

    Private Sub menuDouble_Click(sender As Object, e As RoutedEventArgs) Handles menuDouble.Click
        insertPunctuation(1)
    End Sub

    Private Sub menuLenticular_Click(sender As Object, e As RoutedEventArgs) Handles menuLenticular.Click
        insertPunctuation(2)
    End Sub

    Private Sub menuEllipsis_Click(sender As Object, e As RoutedEventArgs) Handles menuEllipsis.Click
        insertPunctuation(3)
    End Sub

    Private Sub menuWave_Click(sender As Object, e As RoutedEventArgs) Handles menuWave.Click
        insertPunctuation(4)
    End Sub

    Private Sub menuSTitle_Click(sender As Object, e As RoutedEventArgs) Handles menuSTitle.Click
        insertPunctuation(5)
    End Sub

    Private Sub menuDTitle_Click(sender As Object, e As RoutedEventArgs) Handles menuDTitle.Click
        insertPunctuation(6)
    End Sub

    Private Sub menuDash_Click(sender As Object, e As RoutedEventArgs) Handles menuDash.Click
        insertPunctuation(7)
    End Sub

    Private Sub menuClipboard_Click(sender As Object, e As RoutedEventArgs) Handles menuClipboard.Click
        Dim clipboard As String = ""
        Dim count = 0
        Dim clip As Boolean = False

        If tlComplete = 100 Then
            clip = True
        Else
            If MsgBox("Translation Not complete. Are you sure you want to copy translated text to clipboard?", vbYesNo, "Incomplete Translation") = vbYes Then
                clip = True
            Else
                clip = False
            End If
        End If

        If clip = True Then
            While (count <= rawTextArray.GetUpperBound(0))
                If rawTextArray(count) <> "" Then
                    Dim i As Integer = 0
                    While (1)
                        i = i + 1
                        If count + i > rawTextArray.GetUpperBound(0) OrElse Mid(rawTextArray(count + i), 1, 1) <> "$" Then
                            For x As Integer = 1 To i
                                clipboard = clipboard & translatedTextArray(count + (x - 1)) & Space(1)
                            Next
                            clipboard = clipboard & vbNewLine
                            count = count + i
                            Exit While
                        End If
                    End While
                Else
                    clipboard = clipboard & vbNewLine
                    count = count + 1
                End If
            End While
            writeToClipboard(clipboard)
            Task.WaitAll()
            MsgBox("Copy to Clipboard Done", vbOKOnly, "Clipboard")
        End If
    End Sub

    Private Sub menuAbout_Click(sender As Object, e As RoutedEventArgs) Handles menuAbout.Click
        MsgBox("Programmed by: joeglens" & vbNewLine & "Developed Using: Visual Studio Express 2015, Visual Basic," & vbNewLine & "and Windows Presentation Form" & vbNewLine & "Copyright joeglens.wordpress.com 2015" & vbNewLine & "File Version " + My.Application.Info.Version.ToString, vbOKOnly, "About")
    End Sub

    Private Sub currentTranslatedLine_LostFocus(sender As Object, e As RoutedEventArgs) Handles currentTranslatedLine.LostFocus
        caretIndex = currentTranslatedLine.CaretIndex
    End Sub

    Private Sub currentTranslatedLine_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles currentTranslatedLine.PreviewKeyDown
        Dim eof As Boolean

        If rawTextArray IsNot Nothing Then
            If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.End Then
                e.Handled = True
                If currentTranslatedLine.Text <> "" Then
                    tmrCopyToClipboard.IsEnabled = False
                    translatedTextArray(arrayPointer) = currentTranslatedLine.Text
                    calculateProgress()
                    If tlComplete <> 100 Then
                        Dim counter As Integer
                        For counter = 0 To rawTextArray.GetUpperBound(0)
                            If rawTextArray(counter) <> "" AndAlso translatedTextArray(counter) = "" Then
                                Exit For
                            End If
                        Next
                        arrayPointer = counter
                        reviewTop.Text = updateReview(0, arrayPointer - 1)
                        reviewTop.SelectionStart = reviewTop.Text.Length
                        reviewTop.ScrollToEnd()
                        reviewBottom.Text = updateReview(arrayPointer + 1, rawTextArray.GetUpperBound(0))
                        currentRawLine.Text = replaceAndParse(rawTextArray(arrayPointer))
                        currentTranslatedLine.Text = translatedTextArray(arrayPointer)
                        LineStatus.Content = "Line: " & (arrayPointer + 1).ToString & "/" & rawTextArray.Length.ToString
                        Task.WaitAll()
                        tmrCopyToClipboard.IsEnabled = True
                        text2Speech(currentRawLine.Text)
                    End If
                End If
                currentTranslatedLine.Focus()
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.Home Then
                e.Handled = True
                If arrayPointer > 0 Then
                    tmrCopyToClipboard.IsEnabled = False
                    translatedTextArray(arrayPointer) = currentTranslatedLine.Text
                    arrayPointer = 0
                    reviewTop.Text = ""
                    reviewBottom.Text = updateReview(1, rawTextArray.GetUpperBound(0))
                    currentRawLine.Text = replaceAndParse(rawTextArray(0))
                    currentTranslatedLine.Text = translatedTextArray(0)
                    LineStatus.Content = "Line: 1/" & rawTextArray.Length.ToString
                    calculateProgress()
                    Task.WaitAll()
                    tmrCopyToClipboard.IsEnabled = True
                    text2Speech(currentRawLine.Text)
                End If
                currentTranslatedLine.Focus()
            ElseIf e.Key = Key.Enter OrElse e.Key = Key.PageUp OrElse e.Key = Key.PageDown Then
                e.Handled = True
                tmrCopyToClipboard.IsEnabled = False
                e.Handled = True
                eof = False
                translatedTextArray(arrayPointer) = currentTranslatedLine.Text
                If e.Key = Key.Enter OrElse e.Key = Key.PageDown Then
                    While (1)
                        If arrayPointer < rawTextArray.GetUpperBound(0) Then
                            arrayPointer = arrayPointer + 1
                            If arrayPointer > rawTextArray.GetUpperBound(0) Or rawTextArray(arrayPointer) <> "" Then
                                Exit While
                            End If
                        Else
                            eof = True
                            Exit While
                        End If
                    End While
                Else
                    While (1)
                        If arrayPointer = 0 Then
                            eof = True
                            Exit While
                        Else
                            arrayPointer = arrayPointer - 1
                            If rawTextArray(arrayPointer) <> "" Then
                                Exit While
                            End If
                        End If
                    End While
                End If
                If eof = False Then
                    reviewTop.Text = updateReview(0, arrayPointer - 1)
                    reviewTop.SelectionStart = reviewTop.Text.Length
                    reviewTop.ScrollToEnd()
                    reviewBottom.Text = updateReview(arrayPointer + 1, rawTextArray.GetUpperBound(0))
                    currentRawLine.Text = replaceAndParse(rawTextArray(arrayPointer))
                    currentTranslatedLine.Text = translatedTextArray(arrayPointer)
                End If
                LineStatus.Content = "Line: " & (arrayPointer + 1).ToString & "/" & rawTextArray.Length.ToString
                calculateProgress()
                If e.Key = Key.Enter Then
                    translatedTextArray(arrayPointer) = currentTranslatedLine.Text
                    Dim fileout As String = txtOutput
                    For i As Integer = 0 To translatedTextArray.GetUpperBound(0)
                        fileout = fileout & translatedTextArray(i) & vbNewLine
                    Next
                    Dim fileoutput As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(filepath, False)
                    Try
                        fileoutput.Write(fileout)
                        fileoutput.Close()
                    Catch ex As Exception
                        MsgBox("File saving error. Check cuurent file Is Not used by other apps", vbCritical, "File Write Error")
                    End Try
                End If
                Task.WaitAll()
                tmrCopyToClipboard.IsEnabled = True
                currentTranslatedLine.Focus()
                text2Speech(currentRawLine.Text)
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.Right Then
                e.Handled = True
                tmrCopyToClipboard.IsEnabled = False
                If (parsepointer + 1) < parsenumber Then
                    parsepointer = parsepointer + 1
                End If

                If parsepointer = -1 Then
                    currentRawLine.Text = replaceAndParse(rawTextArray(arrayPointer))

                Else
                    currentRawLine.SelectionStart = currentRawLine.Text.IndexOf(parseSentence(parsepointer))
                    currentRawLine.SelectionLength = parseSentence(parsepointer).Length
                End If
                Task.WaitAll()
                tmrCopyToClipboard.IsEnabled = True
                currentTranslatedLine.Focus()
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.Left Then
                e.Handled = True
                tmrCopyToClipboard.IsEnabled = False
                If (parsepointer - 1) >= -1 AndAlso replaced = False Then
                    parsepointer = parsepointer - 1
                ElseIf (parsepointer - 1) >= -2 AndAlso replaced = True Then
                    parsepointer = parsepointer - 1
                End If
                If parsepointer = -1 Then
                    currentRawLine.SelectionLength = 0
                ElseIf parsepointer = -2 Then
                    currentRawLine.Text = rawTextArray(arrayPointer)
                Else
                    currentRawLine.SelectionStart = currentRawLine.Text.IndexOf(parseSentence(parsepointer))
                    currentRawLine.SelectionLength = parseSentence(parsepointer).Length
                End If
                Task.WaitAll()
                tmrCopyToClipboard.IsEnabled = True
                currentTranslatedLine.Focus()
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.S Then
                e.Handled = True
                menuSave_Click(sender, e)
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.L Then
                e.Handled = True
                menuPhrase_Click(sender, e)
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.A Then
                e.Handled = True
                currentTranslatedLine.SelectAll()
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.I Then
                e.Handled = True
                menuClipboard_Click(sender, e)
            ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.J AndAlso currentTranslatedLine.SelectionLength > 0 Then
                e.Handled = True
                Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex", True)
                fileopened.WriteLine(currentTranslatedLine.SelectedText)
                fileopened.Close()
                MsgBox("The word """ & currentTranslatedLine.SelectedText & """ has been added to the dictionary", vbOKOnly, "Dictionary")
                Dim dictionary As IList = SpellCheck.GetCustomDictionaries(currentTranslatedLine)
                dictionary.Clear()
                dictionary.Add(New Uri(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".lex"))
            End If
        End If

        If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.O Then
            e.Handled = True
            menuOpen_Click(sender, e)
        ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.P Then
            e.Handled = True
            menuProfile_Click(sender, e)
        ElseIf e.Key >= Key.F1 AndAlso e.Key <= Key.F8 Then
            e.Handled = True
            insertPunctuation(e.Key - Key.F1)
        ElseIf Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.F Then
            e.Handled = True
            writeToClipboard(currentTranslatedLine.Text)
        End If
    End Sub

    Private Sub reviewTop_KeyDown(sender As Object, e As KeyEventArgs) Handles reviewTop.KeyDown
        currentTranslatedLine_PreviewKeyDown(sender, e)
    End Sub

    Private Sub txtReviewBottom_KeyDown(sender As Object, e As KeyEventArgs) Handles reviewBottom.KeyDown
        currentTranslatedLine_PreviewKeyDown(sender, e)
    End Sub

    Private Sub txtboxCurrentLine_KeyDown(sender As Object, e As KeyEventArgs) Handles currentRawLine.KeyDown
        currentTranslatedLine_PreviewKeyDown(sender, e)
    End Sub

    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        currentTranslatedLine_PreviewKeyDown(sender, e)
    End Sub

    Private Sub currentTranslatedLine_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs) Handles currentTranslatedLine.ContextMenuOpening

        currentTranslatedLine.ContextMenu = New ContextMenu
        Dim caretIndex As Integer = currentTranslatedLine.CaretIndex
        Dim spellError As SpellingError = currentTranslatedLine.GetSpellingError(caretIndex)

        If spellError IsNot Nothing AndAlso spellError.Suggestions.Count > 0 Then
            For Each words As String In spellError.Suggestions
                Dim mi As New MenuItem
                mi.Header = words
                mi.FontWeight = FontWeights.Bold
                mi.Command = EditingCommands.CorrectSpellingError
                mi.CommandParameter = words
                mi.CommandTarget = currentTranslatedLine
                currentTranslatedLine.ContextMenu.Items.Add(mi)
            Next
        Else
            Dim mi2 As New MenuItem
            mi2.Header = "No Spelling Suggestion"
            mi2.FontWeight = FontWeights.Bold
            currentTranslatedLine.ContextMenu.Items.Add(mi2)
        End If
        currentTranslatedLine.ContextMenu.Items.Add(New Separator)
        Dim mi1 As New MenuItem
        mi1.Header = "Add to Dictionary"
        AddHandler mi1.Click, AddressOf AddToDictionary
        mi1.Command = EditingCommands.IgnoreSpellingError
        mi1.CommandTarget = currentTranslatedLine
        If currentTranslatedLine.SelectionLength = 0 Then
            mi1.IsEnabled = False
        End If
        currentTranslatedLine.ContextMenu.Items.Add(mi1)
        currentTranslatedLine.ContextMenu.Items.Add(New Separator)
        Dim mitem As New MenuItem
        mitem.Header = "Copy"
        mitem.Command = ApplicationCommands.Copy
        If currentTranslatedLine.SelectionLength = 0 Then
            mitem.IsEnabled = False
        End If
        currentTranslatedLine.ContextMenu.Items.Add(mitem)
        Dim mitem1 As New MenuItem
        mitem1.Header = "Paste"
        mitem1.Command = ApplicationCommands.Paste
        If Clipboard.ContainsText = False Then
            mitem1.IsEnabled = False
        End If
        currentTranslatedLine.ContextMenu.Items.Add(mitem1)
    End Sub


    Private Sub reviewTop_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles reviewTop.PreviewMouseDoubleClick
        If rawTextArray IsNot Nothing AndAlso arrayPointer > 0 Then
            Dim i As Integer
            i = reviewTop.CaretIndex

            For x As Integer = 0 To arrayPointer
                If i < linenumber(1, x) AndAlso i > linenumber(0, x) Then
                    translatedTextArray(arrayPointer) = currentTranslatedLine.Text
                    arrayPointer = x
                    reviewTop.Text = updateReview(0, arrayPointer - 1)
                    reviewTop.SelectionStart = reviewTop.Text.Length
                    reviewTop.ScrollToEnd()
                    reviewBottom.Text = updateReview(arrayPointer + 1, rawTextArray.GetUpperBound(0))
                    reviewBottom.ScrollToHome()
                    currentRawLine.Text = replaceAndParse(rawTextArray(arrayPointer))
                    currentTranslatedLine.Text = translatedTextArray(arrayPointer)
                    LineStatus.Content = "Line: " & (arrayPointer + 1).ToString & "/" & rawTextArray.Length.ToString
                    calculateProgress()
                    Task.WaitAll()
                    tmrCopyToClipboard.IsEnabled = True
                    text2Speech(currentRawLine.Text)
                    Return
                End If
            Next
            reviewTop.SelectionLength = 0
        End If


    End Sub

    Private Sub reviewBottom_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles reviewBottom.PreviewMouseDoubleClick
        If rawTextArray IsNot Nothing AndAlso arrayPointer < rawTextArray.GetUpperBound(0) Then
            Dim i As Integer
            i = reviewBottom.CaretIndex

            For x As Integer = arrayPointer To rawTextArray.GetUpperBound(0)
                If i < linenumber(1, x) AndAlso i > linenumber(0, x) Then
                    translatedTextArray(arrayPointer) = currentTranslatedLine.Text
                    arrayPointer = x
                    reviewTop.Text = updateReview(0, arrayPointer - 1)
                    reviewTop.SelectionStart = reviewTop.Text.Length
                    reviewTop.ScrollToEnd()
                    reviewBottom.Text = updateReview(arrayPointer + 1, rawTextArray.GetUpperBound(0))
                    reviewBottom.ScrollToHome()
                    currentRawLine.Text = replaceAndParse(rawTextArray(arrayPointer))
                    currentTranslatedLine.Text = translatedTextArray(arrayPointer)
                    LineStatus.Content = "Line: " & (arrayPointer + 1).ToString & "/" & rawTextArray.Length.ToString
                    calculateProgress()
                    Task.WaitAll()
                    tmrCopyToClipboard.IsEnabled = True
                    text2Speech(currentRawLine.Text)
                    Return
                End If
            Next
            reviewBottom.SelectionLength = 0
        End If
    End Sub
End Class
