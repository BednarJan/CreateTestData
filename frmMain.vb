
Imports Microsoft.Office.Interop
Imports DataModel
Imports Devart.Data.SQLite


Public Class frmMain
    Public myExcelFile

    Public Const whoAmI As String = "Jan Bednar"
    Public Const TestStation As String = "HPL1"

    Const rowPosTambient As Integer = 4
    Const rowPosTcoolant As Integer = 3
    Const rowPosCR As Integer = 5
    Const rowPosTestStation As Integer = 6
    Const rowPosTestStrtDate As Integer = 8
    Const rowPosTestFinishDate As Integer = 9
    Const rowPosModel As Integer = 10
    Const rowPosSerialNo As Integer = 11
    Const rowPosHWRevision As Integer = 12
    Const rowPosComponentID As Integer = 13
    Const rowPosCSWID As Integer = 14
    Const rowPosHeaderDef As Integer = 18

    Dim xlApp As Excel.Application = Nothing
    Dim xlBook As Excel.Workbook = Nothing
    Dim xlBooks As Excel.Workbooks = Nothing
    Dim xlSheet As Excel.Worksheet = Nothing
    Dim myFile As String = vbNullString

    Dim myConn As SQLiteConnection = Nothing
    Dim myConnString As String = vbNullString
    Dim AVTTestID As Int64 = -1



    Dim numRows As Long
    Dim iRow As Integer
    Dim tAmb As String = vbNullString
    Dim tCoolant As Decimal = vbNullString



    Private Sub btnOpenExcelFile_Click(sender As Object, e As EventArgs) Handles btnOpenExcelFile.Click
        Dim myTxt As String = vbNullString
        Dim tdsAVTTest As DataModel.dsAVTTest = New dsAVTTest
        Dim drAVTTest As DataRow = Nothing

        Try
            xlApp = New Excel.Application()
            xlBooks = xlApp.Workbooks

            myFile = AppDomain.CurrentDomain.BaseDirectory.Replace("\bin\Debug\", "\ExcelSheets")
            myFile &= "\" & "BCL25-700-8_014_1004289500018_MeasCh3f_CR000223_63Hz_01.xls"

            xlBook = xlBooks.Open(myFile)

            myConnString = "Data Source=D:\Projects\SQLlite_experiments\DB\DB_Files\BCL25-700.db"

            myConn = New SQLiteConnection(myConnString)

            'looping through the sheets
            For Each xlSheet In xlBook.Sheets

                numRows = xlSheet.UsedRange.Rows.Count

                drAVTTest = tdsAVTTest.Tables(0).NewRow


                For iRow = rowPosTambient To rowPosHeaderDef + 1
                    'loop  through thew header 
                    myTxt = xlSheet.Rows(iRow).cells(2).text

                    If xlSheet.Name.Contains("Ref") Then
                        'Reference sheet for the creation header descriptions
                        'will be created just once

                        Select Case iRow


                            Case rowPosCR
                                drAVTTest("CR") = myTxt
                            Case rowPosTestStation
                                drAVTTest("TestStation") = myTxt
                            Case rowPosTestStrtDate
                                drAVTTest("TestStart") = CType(myTxt, Date)
                            Case rowPosTestFinishDate
                                drAVTTest("TestFinish") = CType(myTxt, Date)
                            Case rowPosModel
                                drAVTTest("UUTModel") = myTxt
                            Case rowPosSerialNo
                                drAVTTest("SerialNo") = myTxt
                            Case rowPosHWRevision
                                drAVTTest("HWRevision") = myTxt
                            Case rowPosComponentID
                                drAVTTest("ComponentID") = myTxt
                            Case rowPosCSWID
                                drAVTTest("SWID") = myTxt
                        End Select
                    End If

                    Select Case iRow
                        Case rowPosTambient
                            tAmb = myTxt
                        Case rowPosTcoolant
                            tCoolant = myTxt
                    End Select

                    drAVTTest("TestEngName") = whoAmI
                    drAVTTest("TestStation") = TestStation

                Next


            Next

            myConn.Open()

            AVTTestID = DataModel.ManageAVTTest.AddNewRecord(drAVTTest, myConn)

            'Dim dlgRes As DialogResult = ofDlg.ShowDialog()
            'If dlgRes = DialogResult.OK Then
            '    myFilename = ofDlg.FileName

            '    If IO.File.Exists(myFilename) Then
            '        xlBook = xlBooks.Open(myFilename)

            '        'looping through the sheets
            '        For Each xlSheet In xlBook.Sheets





            '        Next


            '    End If
            'End If




        Catch ex As Exception
        Finally
            myConn.Close()
            xlBook.Close()
            xlApp.Quit()

        End Try

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim path As String = AppDomain.CurrentDomain.BaseDirectory
        path = path.Replace("\bin\Debug\", "\ExcelSheets")
        Me.ofDlg.InitialDirectory = path
    End Sub









End Class
