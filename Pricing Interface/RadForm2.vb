Public Class RadForm2
    Dim oForm1 As RadForm3
    Dim oForm2 As RadAboutBox1
    Dim oForm3 As RadForm2

    Protected Property MainForm As RadForm1

    Public Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click

        If TextBox3.Text = "Test" And TextBox2.Text = "Test" Then


            oForm1 = New RadForm3()

            oForm1.Show()


        Else
            MsgBox("Identification impossible, veuillez réessayer !", MsgBoxStyle.Critical, "Erreur")

        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        oForm2 = New RadAboutBox1()
        oForm2.Show()
    End Sub
End Class
