Public Class Form2
    Dim oForm As RadForm1

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click

        If TextBox3.Text = "Test" And TextBox2.Text = "Test" Then
            oForm = New RadForm1()
            oForm.Show()
        Else
            MsgBox("Identification impossible, veuillez réessayer !", MsgBoxStyle.Critical, "Erreur")

        End If


    End Sub



    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class