Public Class frmPhrase
    Dim textOutput As String = ""
    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        Dim fileopened As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".csv", True)
        fileopened.Write(txtPhrase.Text & "," & Replace(txtTranslation.Text, " ", "_") & vbNewLine)
        fileopened.Close()
        Dim fileread As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(System.AppDomain.CurrentDomain.BaseDirectory & "\Profile\" & My.Settings.ProfileUsed & ".csv")
        textOutput = fileread.ReadToEnd
        fileread.Close()
        DialogResult = True
    End Sub

    Public ReadOnly Property glossaryText As String
        Get
            Return textOutput
        End Get
    End Property

    Private Sub txtTranslation_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtTranslation.TextChanged
        If txtTranslation.Text.Length > 0 Then
            btnOK.IsEnabled = True
        Else
            btnOK.IsEnabled = False
        End If
    End Sub
End Class
