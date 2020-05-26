Public Class frmNew
    Dim rawOutput As String = ""
    Dim filepath As String = ""

    Private Sub btnCreate_Click(sender As Object, e As RoutedEventArgs) Handles btnCreate.Click
        Dim rawTextArray As List(Of String)
        Dim rawParseTextArray As List(Of String)
        Dim blankTranslate As String
        Dim senCount As Integer

        rawTextArray = entryBox.Text.Split(vbCr).ToList
        For i As Integer = 0 To rawTextArray.Count - 1
            rawTextArray(i) = Replace(rawTextArray(i), vbLf, "")
        Next
        rawOutput = ""
        blankTranslate = ""
        For i As Integer = 0 To rawTextArray.Count - 1
            senCount = Split(rawTextArray(i), "。").GetUpperBound(0)
            If senCount < 1 Then
                rawOutput = rawOutput & "%" & rawTextArray(i) & vbNewLine
                blankTranslate = blankTranslate & "" & vbNewLine
            Else
                rawParseTextArray = Split(rawTextArray(i), "。").ToList
                rawParseTextArray.RemoveAll(Function(str) String.IsNullOrEmpty(str))
                rawOutput = rawOutput & "%" & rawParseTextArray(0) & "。" & vbNewLine
                blankTranslate = blankTranslate & "" & vbNewLine
                For x As Integer = 1 To (rawParseTextArray.Count - 1)
                    Dim sentenceEnder As String = ""
                    If (x < rawParseTextArray.Count - 1) OrElse (x = rawParseTextArray.Count - 1 AndAlso Mid(rawTextArray(i), rawTextArray(i).Length, 1) = "。") Then
                        sentenceEnder = "。"
                    End If
                    rawOutput = rawOutput & "$" & rawParseTextArray(x) & sentenceEnder & vbNewLine
                    blankTranslate = blankTranslate & "" & vbNewLine
                Next
                rawParseTextArray.Clear()
            End If
        Next

        rawOutput = rawOutput & vbNewLine & "---SEPERATOR---" & vbNewLine & blankTranslate

        Dim sfdcreate As New Microsoft.Win32.SaveFileDialog()
        sfdcreate.Filter = "Text Files (*.txt)|*.txt"

        If sfdcreate.ShowDialog = True AndAlso sfdcreate.FileName <> "" Then
            Dim fileoutput As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(sfdcreate.FileName, False)
            fileoutput.Write(rawOutput)
            fileoutput.Close()
            filepath = sfdcreate.FileName
            DialogResult = True
        Else
            DialogResult = False
        End If

    End Sub

    Public ReadOnly Property rawOutputText As String
        Get
            Return rawOutput
        End Get
    End Property

    Public ReadOnly Property filename As String
        Get
            Return filepath
        End Get
    End Property

    Private Sub frmNew_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        entryBox.Focus()
    End Sub
End Class
