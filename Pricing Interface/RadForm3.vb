Public Class RadForm3
    Dim oForm1 As RadForm1
    Dim oForm2 As RadAboutBox1
    Dim oForm3 As RadForm2
    Dim oForm4 As RadForm5
    Dim oForm5 As RadForm6
    Protected Property MainForm As RadForm1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        oForm2 = New RadAboutBox1()
        oForm2.Show()
    End Sub

    Private Sub RadButton2_Click(sender As Object, e As EventArgs) Handles RadButton2.Click
        oForm1 = New RadForm1()

        oForm1.Show()

    End Sub

    Private Sub RadButton3_Click(sender As Object, e As EventArgs) Handles RadButton3.Click
        oForm4 = New RadForm5()

        oForm4.Show()

    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click
        oForm5 = New RadForm6()

        oForm5.Show()
    End Sub
End Class

