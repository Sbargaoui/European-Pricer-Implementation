'Imports Access = Microsoft.Office.Interop.Access
Imports Excel = Microsoft.Office.Interop.Excel
'using Microsoft.Office.Interop.Excel

'Imports Word = Microsoft.Office.Interop.Word
'Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
'Imports Telerik.QuickStart.WinControls
'Imports Telerik.WinControls.UI
'Imports Telerik.Charting
'Imports Telerik.WinControls

'Namespace Telerik.Examples.WinControls.ChartView.FirstLook



Public Class Form1

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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'RadMultiColumnComboBox1.DropDownStyle = ComboBoxStyle.DropDownList

        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        Dim a As String() = {"Excel"}

        'RadMultiColumnComboBox1.


        'RadMultiColumnComboBox1.SelectedIndex = 0


        ComboBox1.Items.AddRange(a)
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click

        Select Case ComboBox1.SelectedItem

            Case "Excel"

                Dim oExcel As Excel.Application
                Dim oBook As Excel.Workbook
                Dim oBooks As Excel.Workbooks
                'Dim oSheets As Excel.Worksheets
                Dim oSheet, oSheet2 As Excel.Worksheet

                Dim x1 As String
                'Dim x2 As Double
                Dim x2 As String

                'Start Excel and open the workbook.
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = True
                oBooks = oExcel.Workbooks
                oBook = oBooks.Open("C:\Users\Rim\Documents\TestActif.xlsm")
                'oBook = oBooks.Add("C:\Users\Rim\Documents\TestActif.xlsm")
                oSheet = oBook.Worksheets("Feuil1")
                'oSheet = oBook.Worksheets.Add("sami")
                oSheet2 = oBook.Worksheets("Feuil2")



                'Run the macros.
                'oExcel.Run("DoKbTest")
                'oExcel.Run("DoKbTestWithParameter", "Hello mitch")
                oExcel.Run("affichage")
                oExcel.Run("proc_extraction_cours_dollar")
                oExcel.Run("proc_extraction_cours_euro")
                'oExcel.Run("UserFormDisplay")
                x1 = oSheet.Range("A9").Text
                'x2 = oSheet.Range("B1").Value
                'x2 = x2 + 335
                x2 = oSheet2.Range("A1").Text
                oSheet.Cells(2, 2).Value = "ecriture"
                MsgBox(x1)
                MsgBox(x2)


                'Clean-up: Close the workbook and quit Excel.
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




        End Select
        GC.Collect()


    End Sub



    
End Class
