Public Class frmProfile
    Dim glossaryList As List(Of String)
    Dim dictFilename As String
    Dim parsers As String

    Public ReadOnly Property textOutput As List(Of String)
        Get
            Return glossaryList
        End Get
    End Property

    Private Function refreshDataGrid(ByVal value As List(Of String)) As System.Data.DataTable

        Dim glossaryTable As New System.Data.DataTable
        glossaryTable.Columns.Add(New System.Data.DataColumn("Phrase"))
        glossaryTable.Columns.Add(New System.Data.DataColumn("Translation"))

        For Each entry As String In value
            Dim entryRow As System.Data.DataRow = glossaryTable.NewRow
            Dim phrase() As String = Split(entry, ",", 2)
            entryRow(0) = phrase(0)
            If phrase.GetUpperBound(0) = 1 Then
                entryRow(1) = phrase(1)
            Else
                entryRow(1) = ""
            End If
            glossaryTable.Rows.Add(entryRow)
        Next
        Return glossaryTable
    End Function

    Private Sub frmProfile_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim filename As String
        parseCharactersInput.Text = My.Settings.ParseChar
        For Each foundfile As String In My.Computer.FileSystem.GetFiles(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile", FileIO.SearchOption.SearchTopLevelOnly, {"*.csv"})
            filename = Mid(foundfile, InStrRev(foundfile, "\") + 1)
            comboList.Items.Add(Mid(filename, 1, filename.Length - 4))
        Next

        If comboList.Items.Count > 0 Then
            For i As Integer = 0 To comboList.Items.Count - 1
                If comboList.Items(i) = My.Settings.ProfileUsed Then
                    comboList.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

    End Sub

    Private Sub comboList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles comboList.SelectionChanged
        If comboList.SelectedIndex > -1 Then
            Dim fileOpened As New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & comboList.SelectedItem.ToString & ".csv")
            glossaryList = fileOpened.ReadToEnd.Split(vbNewLine).ToList
            fileOpened.Close()

            For i As Integer = 0 To glossaryList.Count - 1
                glossaryList(i) = Replace(glossaryList(i), vbLf, "")
            Next
            glossaryList.RemoveAll(Function(str) String.IsNullOrEmpty(str))


            dgvPhrases.ItemsSource = refreshDataGrid(glossaryList).DefaultView

            If comboList.SelectedItem = "Default" Then
                btnDelete.IsEnabled = False
            Else
                btnDelete.IsEnabled = True
            End If
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim txtOutput As String = ""
        My.Settings.ParseChar = parseCharactersInput.Text
        My.Settings.ProfileUsed = comboList.SelectedItem
        My.Settings.Save()
        glossaryList.Clear()
        If dgvPhrases.Items.Count > 1 Then
            For i As Integer = 0 To dgvPhrases.Items.Count - 2
                Dim row As System.Data.DataRowView = dgvPhrases.Items(i)
                Dim txtline As String = row(0).ToString & "," & Replace(row(1).ToString, " ", "_")
                txtOutput = txtOutput & txtline & vbNewLine
                glossaryList.Add(txtline)
            Next
            Dim fileOutput As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & comboList.SelectedItem.ToString & ".csv", False)
            fileOutput.Write(txtOutput)
            fileOutput.Close()
            dictFilename = comboList.SelectedItem.ToString
            DialogResult = True
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As RoutedEventArgs) Handles btnCreate.Click
        Dim newWindow As New frmProfileName
        If newWindow.ShowDialog = True Then
            Dim filename As String = newWindow.filename
            Dim fullpath As String = System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & filename
            If My.Computer.FileSystem.FileExists(fullpath) Then
                If MsgBox("File exist, overwrite profile?" & vbNewLine & "Warning: Overwriting profile will erase all data", vbYesNo, "Overwrite") = vbOK Then
                    Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fullpath & ".csv", False)
                    fileopened.WriteLine("")
                    fileopened.Close()
                    fileopened = My.Computer.FileSystem.OpenTextFileWriter(fullpath & ".lex", False)
                    fileopened.WriteLine("#LID 1033")
                    fileopened.Close()
                    comboList.Items.Add(filename)
                    comboList.SelectedIndex = comboList.Items.Count - 1
                End If
            Else
                Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fullpath & ".csv", False)
                fileopened.WriteLine("")
                fileopened.Close()
                fileopened = My.Computer.FileSystem.OpenTextFileWriter(fullpath & ".lex", False)
                fileopened.WriteLine("#LID 1033")
                fileopened.Close()
                comboList.Items.Add(filename)
                comboList.SelectedIndex = comboList.Items.Count - 1
            End If
        End If
    End Sub


    Private Sub btnDelete_Click(sender As Object, e As RoutedEventArgs) Handles btnDelete.Click
        Dim filename As String = comboList.SelectedItem
        If MsgBox("Are you sure you want to delete " & filename & "?", vbYesNo, "Profile Delete") = vbYes Then
            My.Computer.FileSystem.DeleteFile(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & filename & ".csv")
            My.Computer.FileSystem.DeleteFile(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & filename & ".lex")
            Task.WaitAll()
            dgvPhrases.ItemsSource = Nothing
            comboList.Items.Remove(filename)
            comboList.SelectedIndex = 0
            My.Settings.ProfileUsed = comboList.SelectedItem
        End If
    End Sub


    Private Sub dgvPhrases_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgvPhrases.MouseDoubleClick
        Dim selectedValue As String = ""
        Dim row As System.Data.DataRowView = dgvPhrases.Items(dgvPhrases.SelectedIndex)
        selectedValue = row(0).ToString & " = " & row(1).ToString
        If MsgBox("Are you sure you want to delete: " & selectedValue, vbYesNo, "Delete Phrase") = vbYes Then
            glossaryList.RemoveAt(dgvPhrases.SelectedIndex)
            dgvPhrases.ItemsSource = refreshDataGrid(glossaryList).DefaultView
        End If
    End Sub


End Class

