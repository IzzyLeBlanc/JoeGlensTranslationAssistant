Public Class frmProfileName
    Dim textOutput As String = ""

    Public ReadOnly Property filename As String
        Get
            Return textOutput
        End Get
    End Property

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        DialogResult = False
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        Dim filename As String = txtName.Text
        Dim check As String = "\/*:?""<>|"
        For i As Integer = 1 To 10
            filename = Replace(filename, Mid(check, i, 1), "_")
        Next
        textOutput = filename
        DialogResult = True
    End Sub

    Private Sub txtName_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtName.TextChanged
        If txtName.Text.Length > 0 Then
            btnOK.IsEnabled = True
        Else
            btnOK.IsEnabled = False
        End If
    End Sub
End Class
