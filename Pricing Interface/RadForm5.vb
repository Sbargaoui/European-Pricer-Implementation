Imports Excel = Microsoft.Office.Interop.Excel
Imports System
Imports System.IO
Imports Telerik.WinControls.UI
Imports Telerik.Charting

Public Class RadForm5



    Private Sub RadForm5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadForm2.Close()
        'RadForm3.Close()
        Dim a As String() = {"USD/TND", "EUR/TND"}
        Dim c As String() = {"Volatilité historique", "Volatilité paramètrique", "Volatilité implicite"}
        Dim f As String() = {"Call", "Put"}

        ComboBox4.Hide()

        ComboBox1.Items.AddRange(f)
        ComboBox1.SelectedIndex = 0

        'Select 
        ComboBox2.Items.AddRange(a)
        ComboBox2.SelectedIndex = 0

        ComboBox3.Items.AddRange(c)
        ComboBox3.SelectedIndex = 0

        'TextBox1.Hide()
        'extBox2.Hide()
        RadChartView1.Hide()

        



    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click

        PictureBox1.Image = Nothing

        'TextBox1.Clear()
        'TextBox2.Clear()
        RadChartView1.Series.Clear()
        RadChartView1.Hide()

        Dim t2 As Date



        t2 = RadDateTimePicker1.Value.Date
        MsgBox(t2)


        Select Case ComboBox2.SelectedItem

            Case "USD/TND"

                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet, oSheet3, oSheet4 As Excel.Worksheet
                'Dim t As Double

                'Dim x1 As String
                'Dim x2 As Double
                'Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                '()

                oBook = oBooks.Open("C:\Users\Rim\Desktop\projet\InterfaceBackTest.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("BacktestUSD")
                'oSheet = oBook.Worksheets.Add("sami")

                oSheet3 = oBook.Worksheets("Taux USD")
                oSheet4 = oBook.Worksheets("Taux TND")


                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                'oExcel.Run("affichage")
                'oExcel.Run("proc_extraction_cours_dollar")
                'oExcel.Run("proc_extraction_cours_euro")


                'oSheet.Range("H5").Value = t2
                'MsgBox(oSheet.Range("H5").Value)

                oSheet3.Range("P12").Value = t2


                oExcel.Run("MacroUSD")

                'MsgBox(oSheet3.Range("Q12").Value)


                oSheet4.Range("P12").Value = t2


                oExcel.Run("MacroTND")

                'MsgBox(oSheet4.Range("Q12").Value)


                'maturité 
                oSheet.Range("E2").Value = t2
                oSheet.Range("B10").Value = Convert.ToDouble(TextBox4.Text)

                'spot

                'Vol


                'Spot


                'vol
                oSheet.Range("B7").Value = TextBox8.Text

                'Strike

                'oSheet.Range("G2").Value = TextBox8.Text



                oExcel.Run("Calcul")

                TextBox1.Text = oSheet.Range("D15").Value * 100 'Perte
                TextBox2.Text = oSheet.Range("D16").Value * 100 'Gain


                oExcel.Run("Affiche")

                'PictureBox1.Image = Image.FromFile("C:\Users\Rim\Desktop\projet\GainPerte.png")



                oBook.Save()

                oBook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                oBook = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
                oBooks = Nothing
                oExcel.Quit()
                GC.SuppressFinalize(oExcel)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                oExcel = Nothing


                GC.Collect()

                PictureBox1.Image = Image.FromFile("C:\Users\Rim\Desktop\projet\SpotF1.png")


            Case "EUR/TND"


        End Select

        GC.Collect()


        'PictureBox1.Image = Image.FromFile("C:\Users\Rim\Desktop\projet\SpotF.png")

        RadChartView1.Show()
        TextBox1.Show()
        TextBox2.Show()

        Me.RadChartView1.AreaType = ChartAreaType.Pie
        Me.RadChartView1.ShowLegend = True
        'Me.RadChartView1.Title = "Gain/Perte de la stratégie"
        Me.RadChartView1.ShowTitle = True
        Me.RadChartView1.BackColor = Color.Transparent

        Dim series As New PieSeries()
        series.DataPoints.Add(New PieDataPoint(TextBox1.Text, "Perte"))
        series.DataPoints.Add(New PieDataPoint(TextBox2.Text, "Gain"))



        
        series.ShowLabels = True
        series.LabelMode = True

        Me.RadChartView1.Series.Add(series)
    End Sub
End Class

