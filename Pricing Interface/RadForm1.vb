'Imports Access = Microsoft.Office.Interop.Access
Imports Excel = Microsoft.Office.Interop.Excel
Imports System
Imports System.IO

'using Microsoft.Office.Interop.Excel

'Imports Word = Microsoft.Office.Interop.Word
'Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
'Imports Telerik.QuickStart.WinControls
'Imports Telerik.WinControls.UI
'Imports Telerik.Charting
'Imports Telerik.WinControls

'Namespace Telerik.Examples.WinControls.ChartView.FirstLook



Public Class RadForm1

    'Imports Access = Microsoft.Office.Interop.Access

    'Imports Word = Microsoft.Office.Interop.Word
    'Imports PowerPoint = Microsoft.Office.Interop.PowerPoint


    Private Sub Button1_Click(sender As Object, e As EventArgs)

        'Select Case ComboBox1.SelectedItem

        ' Case "Excel"

        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oBooks As Excel.Workbooks

        'Start Excel and open the workbook.
        oExcel = CreateObject("Excel.Application")
        oExcel.Visible = True
        oBooks = oExcel.Workbooks
        oBook = oBooks.Open("C:\Users\Rim\Documents\TestActif.xlsm")

        'Run the macros.
        'oExcel.Run("DoKbTest")
        'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
        oExcel.Run("proc_extraction_cours_dollar")
        oExcel.Run("proc_extraction_cours_euro")

        'Clean-up: Close the workbook and quit Excel.
        oBook.Close(False)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
        oBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
        oBooks = Nothing
        oExcel.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
        oExcel = Nothing


        'End Select
        GC.Collect()
    End Sub
    Dim path As String = Directory.GetCurrentDirectory()
    Dim chemin As String = "C:\Users\Rim\Documents\Sami\Pfe"



    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'RadMultiColumnComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        RadForm2.Close()
        'MsgBox(path)
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        Dim a As String() = {"USD/TND", "EUR/TND"}
        Dim b As String() = {"Modèle de Garman-Kohlhagen", "Méthode des arbres binomiaux", "Méthode des arbres trinomiaux", "Modèle de Merton", "Méthode de Monte-Carlo"}
        Dim c As String() = {"Volatilité historique", "Volatilité paramètrique", "Volatilité implicite"}

        Dim f As String() = {"Call", "Put"}
        'Dim f As String() = {"Call", "Put", "Call Participatif", "Tunnel", "Barrière"}
        'RadMultiColumnComboBox1.


        'RadMultiColumnComboBox1.SelectedIndex = 0

        ComboBox4.Hide()

        ComboBox1.Items.AddRange(f)
        ComboBox1.SelectedIndex = 0

        'Select 
        ComboBox2.Items.AddRange(b)
        ComboBox2.SelectedIndex = 0

        ComboBox3.Items.AddRange(c)
        ComboBox3.SelectedIndex = 0





        ComboBox5.Items.AddRange(a)
        ComboBox5.SelectedIndex = 0
    End Sub

    Public Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click


        PictureBox1.Image = Nothing



        Dim t3 As System.TimeSpan

        Dim t1, t2 As Date

        t1 = "08/04/2015"

        t2 = RadDateTimePicker1.Value.Date

        t3 = t2.Subtract(t1)
        'Dim t As Double


        MsgBox("Bases de données importées", MsgBoxStyle.OkOnly, "Opération réussie")
        'MsgBox(RadDateTimePicker1.Value.Date)

        Me.RadGroupBox7.Text = ""


        Select Case ComboBox1.SelectedItem

            Case "Call"


                Select Case ComboBox5.SelectedItem

                    Case "USD/TND"

                        Select Case ComboBox2.SelectedItem

                            Case "Modèle de Garman-Kohlhagen"

                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Modèle de Garman-Kohlhagen-USD")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux USD")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("B10").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("B12").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B7").Value = TextBox8.Text

                                'Strike
                                't = TextBox7.Text
                                'MsgBox(t)
                                oSheet.Range("B6").Value = TextBox7.Text

                                'MsgBox(oSheet.Range("B6").Value)

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux local
                                TextBox5.Text = oSheet.Range("B8").Value
                                'taux devise
                                TextBox9.Text = oSheet.Range("B9").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("E5").Value
                                'option %
                                TextBox12.Text = oSheet.Range("B18").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("B19").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("B20").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("B21").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("B22").Value
                                'Rho dev locale
                                TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                TextBox18.Text = oSheet.Range("B24").Value
                                'Omega/Lambda
                                TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                TextBox17.Text = oSheet.Range("B26").Value





                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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
                                'oExcel.Quit()
                                Me.RadGroupBox7.Text = ""


                            Case "Modèle à sauts : Jump Diffusion"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook, oBook2 As Excel.Workbook
                                Dim oBooks, oBooks2 As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                oBooks2 = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\JumpDiffusion.xlsm")
                                oBook2 = oBooks2.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("JumpDiffusion")
                                'oSheet = oBook.Worksheets.Add("sami")
                                'oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook2.Worksheets("Taux USD")
                                oSheet4 = oBook2.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E7").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("E9").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B8").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B6").Value = TextBox7.Text



                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value

                                oSheet.Range("B7").Value = TextBox5.Text - TextBox9.Text

                                'prix option
                                TextBox19.Text = oSheet.Range("E10").Value
                                'option %
                                TextBox12.Text = oSheet.Range("B15").Value
                                'Delta
                                'TextBox10.Text = oSheet.Range("J9").Value
                                'Vega
                                'TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                'TextBox15.Text = oSheet.Range("J10").Value
                                'Theta
                                'TextBox13.Text = oSheet.Range("J12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value
                                TextBox10.Hide()
                                TextBox11.Hide()
                                TextBox15.Hide()
                                TextBox13.Hide()
                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
                                oBook.Save()

                                oBook.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                                oBook = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
                                oBooks = Nothing
                                oBook2.Save()

                                oBook2.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook2)
                                oBook2 = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks2)
                                oBooks2 = Nothing
                                oExcel.Quit()
                                GC.SuppressFinalize(oExcel)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                                oExcel = Nothing

                            Case "Méthode des arbres binomiaux"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Binomial USD")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux USD")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E3").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B7").Value = TextBox4.Text

                                'montant
                                oSheet.Range("F11").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B10").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B8").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("F13").Value
                                'option %
                                TextBox12.Text = oSheet.Range("F14").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("J9").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("J10").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("J12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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

                            Case "Méthode des arbres trinomiaux"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Trinomial USD")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux USD")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("F2").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("C8").Value = TextBox4.Text

                                'montant
                                oSheet.Range("C20").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("C13").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("C9").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()

                                'call
                                oSheet3.Range("F9").Value = "1"
                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("C22").Value
                                'option %
                                TextBox12.Text = oSheet.Range("C15").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("C16").Value
                                'Vega
                                ' TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("C17").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("C18").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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


                        End Select


                    Case "EUR/TND"

                        Select Case ComboBox2.SelectedItem

                            Case "Modèle de Garman-Kohlhagen"
                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Modèle de Garman-Kohlhagen-EUR")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux EUR")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("B10").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("B12").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B7").Value = TextBox8.Text
                                'Strike
                                oSheet.Range("B6").Value = TextBox7.Text


                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux local
                                TextBox5.Text() = (oSheet.Range("B8").Value)
                                'taux devise
                                TextBox9.Text() = (oSheet.Range("B9").Value)


                                'prix option
                                TextBox19.Text = oSheet.Range("E5").Value
                                'option %
                                TextBox12.Text = oSheet.Range("B18").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("B19").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("B20").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("B21").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("B22").Value
                                'Rho dev locale
                                TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                TextBox18.Text = oSheet.Range("B24").Value
                                'Omega/Lambda
                                TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                TextBox17.Text = oSheet.Range("B26").Value





                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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
                                'oExcel.Quit()

                            Case "Modèle à sauts : Jump Diffusion"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook, oBook2 As Excel.Workbook
                                Dim oBooks, oBooks2 As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                oBooks2 = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\JumpDiffusion.xlsm")
                                oBook2 = oBooks2.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("JumpDiffusion")
                                'oSheet = oBook.Worksheets.Add("sami")
                                'oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook2.Worksheets("Taux EUR")
                                oSheet4 = oBook2.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E7").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("E9").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B8").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B6").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value

                                oSheet.Range("B7").Value = TextBox5.Text - TextBox9.Text

                                'prix option
                                TextBox19.Text = oSheet.Range("E10").Value
                                'option %
                                TextBox12.Text = oSheet.Range("B15").Value
                                'Delta
                                'TextBox10.Text = oSheet.Range("J9").Value
                                'Vega
                                'TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                'TextBox15.Text = oSheet.Range("J10").Value
                                'Theta
                                'TextBox13.Text = oSheet.Range("J12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value
                                TextBox10.Hide()
                                TextBox11.Hide()
                                TextBox15.Hide()
                                TextBox13.Hide()
                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
                                oBook.Save()

                                oBook.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                                oBook = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
                                oBooks = Nothing
                                oBook2.Save()

                                oBook2.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook2)
                                oBook2 = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks2)
                                oBooks2 = Nothing
                                oExcel.Quit()
                                GC.SuppressFinalize(oExcel)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                                oExcel = Nothing
                            Case "Méthode des arbres binomiaux"
                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Binomial EUR")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux EUR")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E3").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B7").Value = TextBox4.Text

                                'montant
                                oSheet.Range("F11").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B10").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B8").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("F13").Value
                                'option %
                                TextBox12.Text = oSheet.Range("F14").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("J9").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("J10").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("J12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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

                            Case "Méthode des arbres trinomiaux"
                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Trinomial EUR")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux EUR")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("F2").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("C8").Value = TextBox4.Text

                                'montant
                                oSheet.Range("C20").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("C13").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("C9").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()

                                'call
                                oSheet3.Range("F9").Value = "1"
                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("C22").Value
                                'option %
                                TextBox12.Text = oSheet.Range("C15").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("C16").Value
                                'Vega
                                ' TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("C17").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("C18").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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


                        End Select

                End Select

            Case "Put"
                Select Case ComboBox5.SelectedItem

                    Case "USD/TND"

                        Select Case ComboBox2.SelectedItem

                            Case "Modèle de Garman-Kohlhagen"

                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Modèle de Garman-Kohlhagen-USD")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux USD")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("B10").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("B12").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B7").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B6").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux local
                                TextBox5.Text() = (oSheet.Range("B8").Value)
                                'taux devise
                                TextBox9.Text() = (oSheet.Range("B9").Value)


                                'prix option
                                TextBox19.Text = oSheet.Range("E7").Value
                                'option %
                                TextBox12.Text = oSheet.Range("D18").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("D19").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("D20").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("D21").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("D22").Value
                                'Rho dev locale
                                TextBox14.Text = oSheet.Range("D23").Value
                                'Rho etrangere
                                TextBox18.Text = oSheet.Range("D24").Value
                                'Omega/Lambda
                                TextBox16.Text = oSheet.Range("D25").Value
                                'Vanna
                                TextBox17.Text = oSheet.Range("D26").Value





                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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
                                'oExcel.Quit()
                                Me.RadGroupBox7.Text = ""

                            Case "Modèle à sauts : Jump Diffusion"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook, oBook2 As Excel.Workbook
                                Dim oBooks, oBooks2 As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                oBooks2 = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\JumpDiffusion.xlsm")
                                oBook2 = oBooks2.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("JumpDiffusion")
                                'oSheet = oBook.Worksheets.Add("sami")
                                'oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook2.Worksheets("Taux USD")
                                oSheet4 = oBook2.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E7").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("E9").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B8").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B6").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()

                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value

                                oSheet.Range("B7").Value = TextBox5.Text - TextBox9.Text
                                'prix option
                                TextBox19.Text = oSheet.Range("E11").Value
                                'option %
                                TextBox12.Text = oSheet.Range("B16").Value
                                'Delta
                                'TextBox10.Text = oSheet.Range("J9").Value
                                'Vega
                                'TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                'TextBox15.Text = oSheet.Range("J10").Value
                                'Theta
                                'TextBox13.Text = oSheet.Range("J12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value
                                TextBox10.Hide()
                                TextBox11.Hide()
                                TextBox15.Hide()
                                TextBox13.Hide()
                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
                                oBook.Save()

                                oBook.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                                oBook = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
                                oBooks = Nothing
                                oBook2.Save()

                                oBook2.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook2)
                                oBook2 = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks2)
                                oBooks2 = Nothing
                                oExcel.Quit()
                                GC.SuppressFinalize(oExcel)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                                oExcel = Nothing

                            Case "Méthode des arbres binomiaux"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Binomial USD")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux USD")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E3").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B7").Value = TextBox4.Text

                                'montant
                                oSheet.Range("F11").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B10").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B8").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("F12").Value
                                'option %
                                TextBox12.Text = oSheet.Range("E9").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("K9").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("K11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("K10").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("K12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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

                            Case "Méthode des arbres trinomiaux"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Trinomial USD")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux USD")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("F2").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("C8").Value = TextBox4.Text

                                'montant
                                oSheet.Range("C20").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("C13").Value = TextBox8.Text
                                'Strike
                                oSheet.Range("C9").Value = TextBox7.Text


                                RadGroupBox6.Show()
                                RadGroupBox7.Show()

                                'call
                                oSheet3.Range("F9").Value = "2"
                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("C22").Value
                                'option %
                                TextBox12.Text = oSheet.Range("C15").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("C16").Value
                                'Vega
                                ' TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("C17").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("C18").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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


                        End Select


                    Case "EUR/TND"

                        Select Case ComboBox2.SelectedItem

                            Case "Modèle de Garman-Kohlhagen"
                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Modèle de Garman-Kohlhagen-EUR")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux EUR")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("B10").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("B12").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B7").Value = TextBox8.Text

                                'Strike
                                TextBox7.Text() = oSheet.Range("B28").Value

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux local
                                TextBox5.Text() = (oSheet.Range("B8").Value)
                                'taux devise
                                TextBox9.Text() = (oSheet.Range("B9").Value)


                                'prix option
                                TextBox19.Text = oSheet.Range("E7").Value
                                'option %
                                TextBox12.Text = oSheet.Range("D18").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("D19").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("D20").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("D21").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("D22").Value
                                'Rho dev locale
                                TextBox14.Text = oSheet.Range("D23").Value
                                'Rho etrangere
                                TextBox18.Text = oSheet.Range("D24").Value
                                'Omega/Lambda
                                TextBox16.Text = oSheet.Range("D25").Value
                                'Vanna
                                TextBox17.Text = oSheet.Range("D26").Value





                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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
                                'oExcel.Quit()

                            Case "Modèle à sauts : Jump Diffusion"

                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook, oBook2 As Excel.Workbook
                                Dim oBooks, oBooks2 As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                oBooks2 = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\JumpDiffusion.xlsm")
                                oBook2 = oBooks2.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("JumpDiffusion")
                                'oSheet = oBook.Worksheets.Add("sami")
                                'oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook2.Worksheets("Taux EUR")
                                oSheet4 = oBook2.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E7").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B5").Value = TextBox4.Text

                                'montant
                                oSheet.Range("E9").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B8").Value = TextBox8.Text

                                'Strike
                                oSheet.Range("B6").Value = TextBox7.Text

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()



                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value

                                oSheet.Range("B7").Value = TextBox5.Text - TextBox9.Text

                                'prix option
                                TextBox19.Text = oSheet.Range("E11").Value
                                'option %
                                TextBox12.Text = oSheet.Range("B16").Value
                                'Delta
                                'TextBox10.Text = oSheet.Range("J9").Value
                                'Vega
                                'TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                'TextBox15.Text = oSheet.Range("J10").Value
                                'Theta
                                'TextBox13.Text = oSheet.Range("J12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value
                                TextBox10.Hide()
                                TextBox11.Hide()
                                TextBox15.Hide()
                                TextBox13.Hide()
                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
                                oBook.Save()

                                oBook.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                                oBook = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
                                oBooks = Nothing
                                oBook2.Save()

                                oBook2.Close(False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook2)
                                oBook2 = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks2)
                                oBooks2 = Nothing
                                oExcel.Quit()
                                GC.SuppressFinalize(oExcel)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                                oExcel = Nothing

                            Case "Méthode des arbres binomiaux"
                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Binomial EUR")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux EUR")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("E3").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("B7").Value = TextBox4.Text

                                'montant
                                oSheet.Range("F11").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("B10").Value = TextBox8.Text

                                'Strike
                                TextBox7.Text() = oSheet.Range("B8").Value

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()


                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value


                                'prix option
                                TextBox19.Text = oSheet.Range("F13").Value
                                'option %
                                TextBox12.Text = oSheet.Range("E9").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("K9").Value
                                'Vega
                                TextBox11.Text = oSheet.Range("K11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("K10").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("K12").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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

                            Case "Méthode des arbres trinomiaux"
                                Me.RadGroupBox7.Text = ""
                                Dim oExcel As Excel.Application
                                Dim oBook As Excel.Workbook
                                Dim oBooks As Excel.Workbooks
                                'Dim oSheets As Excel.Worksheets
                                Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

                                'Dim x1 As String
                                'Dim x2 As Double
                                'Dim x2 As String

                                'Start Excel and open the workbook.
                                oExcel = CreateObject("Excel.Application")
                                oExcel.Visible = False
                                oBooks = oExcel.Workbooks
                                '()

                                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                                oSheet = oBook.Worksheets("Arbre Trinomial EUR")
                                'oSheet = oBook.Worksheets.Add("sami")
                                oSheet2 = oBook.Worksheets("SPOTs")
                                oSheet3 = oBook.Worksheets("Taux EUR")
                                oSheet4 = oBook.Worksheets("Taux TND")


                                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                                'oExcel.Run("affichage")
                                'oExcel.Run("proc_extraction_cours_dollar")
                                'oExcel.Run("proc_extraction_cours_euro")

                                oSheet3.Range("P4").Value = t2
                                oSheet4.Range("P4").Value = t2


                                'maturité 
                                oSheet.Range("F2").Value = Convert.ToDouble(t3.Days)

                                'spot
                                oSheet.Range("C8").Value = TextBox4.Text

                                'montant
                                oSheet.Range("C20").Value = TextBox6.Text

                                'Vol
                                oSheet.Range("C13").Value = TextBox8.Text

                                'Strike
                                TextBox7.Text() = oSheet.Range("C9").Value

                                RadGroupBox6.Show()
                                RadGroupBox7.Show()

                                'call
                                oSheet3.Range("F9").Value = "2"
                                'taux devise
                                TextBox9.Text() = oSheet3.Range("Q4").Value / 100
                                'taux local
                                TextBox5.Text() = oSheet4.Range("Q4").Value

                                'prix option
                                TextBox19.Text = oSheet.Range("C22").Value
                                'option %
                                TextBox12.Text = oSheet.Range("C15").Value
                                'Delta
                                TextBox10.Text = oSheet.Range("C16").Value
                                'Vega
                                ' TextBox11.Text = oSheet.Range("J11").Value
                                'Gamma
                                TextBox15.Text = oSheet.Range("C17").Value
                                'Theta
                                TextBox13.Text = oSheet.Range("C18").Value
                                'Rho dev locale
                                'TextBox14.Text = oSheet.Range("B23").Value
                                'Rho etrangere
                                'TextBox18.Text = oSheet.Range("B24").Value
                                ''Omega/Lambda
                                'TextBox16.Text = oSheet.Range("B25").Value
                                'Vanna
                                'TextBox17.Text = oSheet.Range("B26").Value

                                TextBox14.Hide()
                                TextBox16.Hide()
                                TextBox17.Hide()
                                TextBox18.Hide()




                                'Clean-up: Close the workbook and quit Excel.
                                'Application.
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


                        End Select

                End Select

        End Select
        GC.Collect()

        Select Case ComboBox1.SelectedItem

            Case "Call"

                PictureBox1.Image = Nothing

                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet As Excel.Worksheet

                'Dim x1 As String
                'Dim x2 As Double
                'Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                '()

                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\PayoffCall.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("Payoff Values")



                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                'oExcel.Run("affichage")
                'oExcel.Run("proc_extraction_cours_dollar")
                'oExcel.Run("proc_extraction_cours_euro")






                'spot
                oSheet.Range("B2").Value = TextBox4.Text

                'montant
                'oSheet.Range("B12").Value = TextBox6.Text

                'Vol
                'oSheet.Range("B7").Value = TextBox8.Text

                'Strike
                oSheet.Range("B3").Value = TextBox7.Text

                TextBox12.Text = Replace(TextBox12.Text, ",", ".")
                oSheet.Range("D2").Value = TextBox12.Text




                oExcel.Run("Affiche")




                'Clean-up: Close the workbook and quit Excel.
                'Application.
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


                PictureBox1.Image = Image.FromFile("C:\Users\Rim\Documents\Sami\Pfe\call.png")

            Case "Put"

                PictureBox1.Image = Nothing


                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet As Excel.Worksheet

                'Dim x1 As String
                'Dim x2 As Double
                'Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                '()

                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\PayoffPut.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("Payoff Values")



                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                'oExcel.Run("affichage")
                'oExcel.Run("proc_extraction_cours_dollar")
                'oExcel.Run("proc_extraction_cours_euro")






                'spot
                oSheet.Range("B2").Value = TextBox4.Text

                'montant
                'oSheet.Range("B12").Value = TextBox6.Text

                'Vol
                'oSheet.Range("B7").Value = TextBox8.Text

                'Strike
                oSheet.Range("B3").Value = TextBox7.Text


                TextBox12.Text = Replace(TextBox12.Text, ",", ".")
                oSheet.Range("I2").Value = TextBox12.Text

                oExcel.Run("Affiche")




                'Clean-up: Close the workbook and quit Excel.
                'Application.
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


                PictureBox1.Image = Image.FromFile("C:\Users\Rim\Documents\Sami\Pfe\put.png")


        End Select




        'GC.Collect()

    End Sub





    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim d As String() = {"Moyenne Exponentielle Mobile EMWA", "GARCH(1,1)"}

        If ComboBox3.SelectedItem.ToString = "Volatilité paramètrique" Then
            ComboBox4.Items.Clear()
            ComboBox4.Visible = True
            ComboBox4.Items.AddRange(d)
            ComboBox4.SelectedIndex = 0
        Else
            ComboBox4.Hide()
        End If





    End Sub



    



    Private Sub RadButton2_Click(sender As Object, e As EventArgs) Handles RadButton2.Click
        Dim t3 As System.TimeSpan

        Dim t1, t2 As Date

        t1 = "08/04/2015"

        t2 = RadDateTimePicker1.Value.Date

        t3 = t2.Subtract(t1)

        If ComboBox5.SelectedItem.ToString = "USD/TND" Then



            If ComboBox3.SelectedItem.ToString = "Volatilité historique" Then

                ComboBox4.Hide()

                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet As Excel.Worksheet

                'Dim x1 As String
                'Dim x2 As Double
                'Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                '()

                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("SPOTs")
                'oSheet = oBook.Worksheets.Add("sami")


                oSheet.Range("R5").Value = t2



                TextBox8.Text = oSheet.Range("T5").Value

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



            ElseIf ComboBox3.SelectedItem.ToString = "Volatilité paramètrique" Then

                ComboBox4.Visible = True
                'ComboBox4.Items.AddRange(d)
                'ComboBox4.SelectedIndex = 0


                If ComboBox4.SelectedItem.ToString = "GARCH(1,1)" Then

                    Dim oExcel As Excel.Application
                    Dim oBook As Excel.Workbook
                    Dim oBooks As Excel.Workbooks
                    'Dim oSheets As Excel.Worksheets
                    Dim oSheet As Excel.Worksheet

                    'Dim x1 As String
                    'Dim x2 As Double
                    'Dim x2 As String

                    'Start Excel and open the workbook.
                    oExcel = CreateObject("Excel.Application")
                    oExcel.Visible = False
                    oBooks = oExcel.Workbooks
                    '()

                    oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Garch USD.xlsm")
                    'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                    oSheet = oBook.Worksheets("Raw Data")
                    'oSheet = oBook.Worksheets.Add("sami")

                    TextBox8.Text = oSheet.Range("H13").Value

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


                ElseIf ComboBox4.SelectedItem.ToString = "Moyenne Exponentielle Mobile EMWA" Then

                    Dim oExcel As Excel.Application
                    Dim oBook As Excel.Workbook
                    Dim oBooks As Excel.Workbooks
                    'Dim oSheets As Excel.Worksheets
                    Dim oSheet As Excel.Worksheet

                    'Dim x1 As String
                    'Dim x2 As Double
                    'Dim x2 As String

                    'Start Excel and open the workbook.
                    oExcel = CreateObject("Excel.Application")
                    oExcel.Visible = False
                    oBooks = oExcel.Workbooks
                    '()

                    oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Garch USD.xlsm")
                    'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                    oSheet = oBook.Worksheets("Raw Data")
                    'oSheet = oBook.Worksheets.Add("sami")

                    TextBox8.Text = oSheet.Range("H13").Value - oSheet.Range("H12").Value

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




                End If

            ElseIf ComboBox3.SelectedItem.ToString = "Volatilité implicite" Then
                ComboBox4.Hide()
                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet As Excel.Worksheet

                'Dim x1 As String
                'Dim x2 As Double
                'Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                '()

                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("Implicite USD")
                'oSheet = oBook.Worksheets.Add("sami")


                'oSheet.Range("F7").Value = Convert.ToDouble(t3.Days)



                TextBox8.Text = oSheet.Range("E16").Value

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




            End If

        ElseIf ComboBox5.SelectedItem.ToString = "EUR/TND" Then

            If ComboBox3.SelectedItem.ToString = "Volatilité historique" Then

                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet As Excel.Worksheet

                'Dim x1 As String
                'Dim x2 As Double
                'Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                '()

                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("SPOTs")
                'oSheet = oBook.Worksheets.Add("sami")


                oSheet.Range("R5").Value = t2



                TextBox8.Text = oSheet.Range("S5").Value

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


            ElseIf ComboBox3.SelectedItem.ToString = "Volatilité paramètrique" Then

                ComboBox4.Visible = True
                'ComboBox4.Items.AddRange(d)
                'ComboBox4.SelectedIndex = 0


                If ComboBox4.SelectedItem.ToString = "GARCH(1,1)" Then

                    Dim oExcel As Excel.Application
                    Dim oBook As Excel.Workbook
                    Dim oBooks As Excel.Workbooks
                    'Dim oSheets As Excel.Worksheets
                    Dim oSheet As Excel.Worksheet

                    'Dim x1 As String
                    'Dim x2 As Double
                    'Dim x2 As String

                    'Start Excel and open the workbook.
                    oExcel = CreateObject("Excel.Application")
                    oExcel.Visible = False
                    oBooks = oExcel.Workbooks
                    '()

                    oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Garch EUR.xlsm")
                    'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                    oSheet = oBook.Worksheets("Raw Data")
                    'oSheet = oBook.Worksheets.Add("sami")

                    TextBox8.Text = oSheet.Range("H13").Value

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


                ElseIf ComboBox4.SelectedItem.ToString = "Moyenne Exponentielle Mobile EMWA" Then

                    Dim oExcel As Excel.Application
                    Dim oBook As Excel.Workbook
                    Dim oBooks As Excel.Workbooks
                    'Dim oSheets As Excel.Worksheets
                    Dim oSheet As Excel.Worksheet

                    'Dim x1 As String
                    'Dim x2 As Double
                    'Dim x2 As String

                    'Start Excel and open the workbook.
                    oExcel = CreateObject("Excel.Application")
                    oExcel.Visible = False
                    oBooks = oExcel.Workbooks
                    '()

                    oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Garch EUR.xlsm")
                    'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                    oSheet = oBook.Worksheets("Raw Data")
                    'oSheet = oBook.Worksheets.Add("sami")

                    TextBox8.Text = oSheet.Range("H13").Value - oSheet.Range("H12").Value

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


                End If


            ElseIf ComboBox3.SelectedItem.ToString = "Volatilité implicite" Then

                ComboBox4.Hide()
                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet As Excel.Worksheet

                'Dim x1 As String
                'Dim x2 As Double
                'Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                '()

                oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("Implicite EUR")
                'oSheet = oBook.Worksheets.Add("sami")


                'oSheet.Range("G7").Value = Convert.ToDouble(t3.Days)



                TextBox8.Text = oSheet.Range("E16").Value

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



            End If
        End If
    End Sub

    Private Sub RadButton3_Click(sender As Object, e As EventArgs) Handles RadButton3.Click

        Dim t3 As System.TimeSpan

        Dim t1, t2 As Date

        t1 = "08/04/2015"

        t2 = RadDateTimePicker1.Value.Date

        t3 = t2.Subtract(t1)


        If ComboBox5.SelectedItem.ToString = "USD/TND" Then

            Dim oExcel As Excel.Application
            Dim oBook As Excel.Workbook
            Dim oBooks As Excel.Workbooks
            'Dim oSheets As Excel.Worksheets
            Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet
            'Dim t As Double

            'Dim x1 As String
            'Dim x2 As Double
            'Dim x2 As String

            'Start Excel and open the workbook.
            oExcel = CreateObject("Excel.Application")
            oExcel.Visible = False
            oBooks = oExcel.Workbooks
            '()

            oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
            'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
            oSheet = oBook.Worksheets("Modèle de Garman-Kohlhagen-USD")
            'oSheet = oBook.Worksheets.Add("sami")
            oSheet2 = oBook.Worksheets("SPOTs")
            oSheet3 = oBook.Worksheets("Taux USD")
            oSheet4 = oBook.Worksheets("Taux TND")


            'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
            'oExcel.Run("affichage")
            'oExcel.Run("proc_extraction_cours_dollar")
            'oExcel.Run("proc_extraction_cours_euro")

            oSheet3.Range("P4").Value = t2
            oSheet4.Range("P4").Value = t2


            'maturité 
            oSheet.Range("B10").Value = Convert.ToDouble(t3.Days)

            'spot

            'Vol
            If TextBox8.Text = "" Or TextBox4.Text = "" Then

                MsgBox("Volatilité ou Spot manquants", MsgBoxStyle.Exclamation, "Erreur")

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

            ElseIf TextBox8.Text <> "" And TextBox4.Text <> "" Then

                'Spot
                oSheet.Range("B5").Value = TextBox4.Text

                'vol
                oSheet.Range("B7").Value = TextBox8.Text




                TextBox7.Text = Convert.ToString(oSheet.Range("B28").Value)

                'TextBox7.Text.Substring(0, TextBox7.Text.Length - 12)

                TextBox7.Text = Replace(TextBox7.Text, ",", ".")

                'MsgBox(TextBox7.Text)


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
            End If


        ElseIf ComboBox5.SelectedItem.ToString = "EUR/TND" Then

            Dim oExcel As Excel.Application
            Dim oBook As Excel.Workbook
            Dim oBooks As Excel.Workbooks
            'Dim oSheets As Excel.Worksheets
            Dim oSheet, oSheet2, oSheet3, oSheet4 As Excel.Worksheet

            'Dim x1 As String
            'Dim x2 As Double
            'Dim x2 As String

            'Start Excel and open the workbook.
            oExcel = CreateObject("Excel.Application")
            oExcel.Visible = False
            oBooks = oExcel.Workbooks
            '()

            oBook = oBooks.Open("C:\Users\Rim\Documents\Sami\Pfe\Interface.xlsm")
            'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
            oSheet = oBook.Worksheets("Modèle de Garman-Kohlhagen-EUR")
            'oSheet = oBook.Worksheets.Add("sami")
            oSheet2 = oBook.Worksheets("SPOTs")
            oSheet3 = oBook.Worksheets("Taux EUR")
            oSheet4 = oBook.Worksheets("Taux TND")


            'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
            'oExcel.Run("affichage")
            'oExcel.Run("proc_extraction_cours_dollar")
            'oExcel.Run("proc_extraction_cours_euro")

            oSheet3.Range("P4").Value = t2
            oSheet4.Range("P4").Value = t2


            'maturité 
            oSheet.Range("B10").Value = Convert.ToDouble(t3.Days)

            'spot
            'oSheet.Range("B5").Value = TextBox4.Text



            If TextBox8.Text = "" Or TextBox4.Text = "" Then

                MsgBox("Volatilité ou Spot manquants", MsgBoxStyle.Exclamation, "Erreur")

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

            ElseIf TextBox8.Text <> "" And TextBox4.Text <> "" Then


                oSheet.Range("B5").Value = TextBox4.Text

                'vol
                oSheet.Range("B7").Value = TextBox8.Text


                TextBox7.Text = Convert.ToString(oSheet.Range("B28").Value)


                TextBox7.Text = Replace(TextBox7.Text, ",", ".")

                'MsgBox(TextBox7.Text)


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
            End If

        End If






    End Sub
End Class
