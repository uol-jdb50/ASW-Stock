Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core

Public Class Form1
    Private con As OleDb.OleDbConnection
    Public DBLocation As String
    Public ReportLocation As String
    Public TemplateLocation As String
    Public numLargeBoxIn As Integer
    Public numLargeBoxStore As Integer
    Public invDate As Date
    'Private cmd As New OleDbCommand
    Public query As String = "SELECT * FROM Goods"
    Public DeliveryInbound As Decimal
    Public StoragePerWeek As Decimal
    Public IncomingKitCheck As Decimal
    Public LargeBoxInbound As Decimal
    Public LargeBoxStoragePerWeek As Decimal
    Public PickPackFeePerItem As Decimal
    Public ManagementFeePerMonth As Decimal
    Public DespatchofKit As Decimal
    Public HourlyRate As Decimal
    Public invTotal As Decimal
    Public DespatchTotal As Decimal
    Public CurrentSpareDate As Date

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        invDate = DateTimePicker1.Value
        invDate = invDate.AddDays(Date.DaysInMonth(invDate.Year, invDate.Month) - invDate.Day)

        If DateSerial(Year(invDate), Month(invDate), 0) <> CurrentSpareDate Then
            ToolStripStatusLabel1.Text = "Invoice processing failed: Run spare days reset first."
            Exit Sub
        End If

        DBLocation = txtDB.Text
        TemplateLocation = txtReports.Text & "\"
        ReportLocation = TemplateLocation
        'Do
        '    Dim c As Char = Strings.Right(ReportLocation, 1)
        '    If c = "/" Or c = "\" Then
        '        Exit Do
        '    Else
        '        ReportLocation = Strings.Left(ReportLocation, ReportLocation.Length - 1)
        '    End If
        'Loop While (True)

        con = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBLocation & ";Persist Security Info=False;")


        Dim prices As New OleDbCommand("SELECT * FROM Charges", con)

        Using con
            Try
                con.Open()
            Catch ex As Exception
                ToolStripStatusLabel1.Text = "Error: Could not find database."
                Exit Sub
            End Try

            Dim reader As OleDbDataReader = prices.ExecuteReader()

            If reader.HasRows Then
                Do While reader.Read()
                    DeliveryInbound = reader.GetDecimal(0)
                    StoragePerWeek = reader.GetDecimal(1)
                    IncomingKitCheck = reader.GetDecimal(2)
                    LargeBoxInbound = reader.GetDecimal(3)
                    LargeBoxStoragePerWeek = reader.GetDecimal(4)
                    PickPackFeePerItem = reader.GetDecimal(5)
                    ManagementFeePerMonth = reader.GetDecimal(6)
                    DespatchofKit = reader.GetDecimal(7)
                    HourlyRate = reader.GetDecimal(8)
                Loop
            End If
        End Using

        con = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBLocation & ";Persist Security Info=False;")

        Dim cmd As New OleDbCommand("SELECT * FROM Goods", con)

        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        Dim recordCount As Integer = 1

        oExcel = CreateObject("Excel.Application")
        oExcel.Visible = False
        oBook = oExcel.Workbooks.Add
        oSheet = oBook.Worksheets(1)

        oSheet.Range("A" & recordCount).Value = "Serial No."
        oSheet.Range("B" & recordCount).Value = "Delivery Inbound"
        oSheet.Range("C" & recordCount).Value = "Storage"
        oSheet.Range("D" & recordCount).Value = "Kit Check"
        oSheet.Range("E" & recordCount).Value = "Pick Fee"
        oSheet.Range("F" & recordCount).Value = "Dispatch"
        oSheet.Range("G" & recordCount).Value = "Total"
        recordCount = recordCount + 1
        Using con

            con.Open()

            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            If reader.HasRows Then
                Do While reader.Read()
                    Dim SerialNumber As Integer = 0
                    Dim DateIn As Date = Nothing
                    Dim DateOut As Date = Nothing
                    Dim FaultyReturn As Boolean = False
                    Dim IncomingCheck As Boolean = False
                    Dim SpareDays As Integer = 0
                    Dim Total As Decimal = 0

                    SerialNumber = reader.GetInt32(0)
                    ToolStripStatusLabel1.Text = "Processing serial number " & SerialNumber
                    Try
                        DateIn = reader.GetDateTime(2)
                    Catch ex As InvalidCastException
                        DateIn = Nothing
                    End Try
                    Try
                        DateOut = reader.GetDateTime(3)
                    Catch ex As InvalidCastException
                        DateOut = Nothing
                    End Try
                    FaultyReturn = reader.GetBoolean(6)
                    IncomingCheck = reader.GetBoolean(8)
                    Try
                        If invDate.Month = DateSerial(Year(Today), Month(Today), 0).Month Then
                            SpareDays = reader.GetInt32(9)
                        End If

                    Catch ex As InvalidCastException
                        SpareDays = 0
                    End Try
                    '(DateIn <> Nothing And DateIn <= invDate) And (DateOut = Nothing Or DateOut > invDate Or (DateOut < invDate And DateOut.Month = invDate.Month And DateOut.Year = invDate.Year))
                    If (DateIn <> Nothing And DateIn <= invDate) And (DateOut = Nothing Or DateOut >= invDate Or (DateOut.Month = invDate.Month And DateOut.Year = invDate.Year)) Then 'CONFIRMED CORRECT
                        'If (DateIn <> Nothing And DateIn <= invDate) Then
                        oSheet.Range("A" & recordCount).Value = SerialNumber
                        'Delivery inbound --------------- CONFIRMED ------------------
                        If (DateIn.Month = invDate.Month) And (DateIn.Year = invDate.Year) Then
                            oSheet.Range("B" & recordCount).Value = FormatNumber(CDbl(DeliveryInbound), 2)
                            'Total += DeliveryInbound
                        End If
                        'Storage --------------- CONFIRMED --------------------
                        If (DateIn <> Nothing And DateIn <= invDate) And (DateOut = Nothing Or DateOut >= DateSerial(Year(invDate), Month(invDate), 1)) Then
                            If (DateIn.Month = DateOut.Month) And (DateIn.Year = DateOut.Year) And (DateIn.Month = invDate.Month) And (DateIn.Year = invDate.Year) Then
                                oSheet.Range("C" & recordCount).Value = FormatNumber(CDbl((((invDate.Day - DateIn.Day + 1) \ 7) * StoragePerWeek)), 2)
                                'SpareDays = (invDate.Day - DateIn.Day + 1 + SpareDays) Mod 7
                            ElseIf (DateOut.Month = invDate.Month) And (DateOut.Year = invDate.Year) Then
                                oSheet.Range("C" & recordCount).Value = FormatNumber(CDbl((((DateOut.Day) \ 7) * StoragePerWeek)), 2)
                                'SpareDays = (DateOut.Day + SpareDays) Mod 7
                            ElseIf (DateIn.Month = invDate.Month) And (DateIn.Year = invDate.Year) Then
                                oSheet.Range("C" & recordCount).Value = FormatNumber(CDbl((((invDate.Day - DateIn.Day + 1) \ 7) * StoragePerWeek)), 2)
                                'SpareDays = (DateOut.Day - DateIn.Day + 1 + SpareDays) Mod 7
                            Else
                                oSheet.Range("C" & recordCount).Value = FormatNumber(CDbl((((invDate.Day) \ 7) * StoragePerWeek)), 2)
                                'SpareDays = (invDate.Day + SpareDays) Mod 7
                            End If
                        End If

                        'Kit check ----------------- CONFIRMED ---------------------
                        If DateIn.Month = invDate.Month And DateIn.Year = invDate.Year Then
                            If IncomingCheck = True Then
                                oSheet.Range("D" & recordCount).Value = FormatNumber(CDbl(IncomingKitCheck), 2)
                            End If
                        End If
                        'Pick pack and despatch
                        If DateOut.Month = invDate.Month And DateOut.Year = invDate.Year Then
                            oSheet.Range("E" & recordCount).Value = FormatNumber(CDbl(PickPackFeePerItem), 2)
                            oSheet.Range("F" & recordCount).Value = FormatNumber(CDbl(DespatchofKit), 2)
                        End If

                        oSheet.Range("G" & recordCount).Value = "=SUM(B" & recordCount & ":F" & recordCount & ")"

                        'UPDATE QUERY RUNS HERE
                        'If invDate.Month = DateSerial(Year(Today), Month(Today), 0).Month Then
                        '    Dim update As String = "UPDATE Goods SET SpareDays = " & SpareDays & " WHERE SerialNumber = " & SerialNumber & ""
                        '    Dim updateCmd As New OleDbCommand(update, con)

                        '    updateCmd.ExecuteNonQuery()
                        'End If
                        recordCount = recordCount + 1
                    End If 'End of record processing

                Loop

            End If

            oSheet.Range("B2:G" & recordCount).NumberFormat = "£###,###,##0.00"

            'oDoc.Content.InsertAfter(vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "DESPATCH TOTAL" & vbTab & "£" & FormatNumber(CDbl(DespatchTotal), 2) & vbCrLf)
            'oDoc.Content.InsertAfter(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SUBTOTAL" & vbTab & vbTab & "£" & FormatNumber(CDbl(invTotal), 2))
            oSheet.Range("E" & recordCount + 1).Value = "Despatch Total"
            oSheet.Range("G" & recordCount + 1).Value = "=SUM(F2:F" & recordCount & ")"
            oSheet.Range("E" & recordCount + 2).Value = "Subtotal"
            oSheet.Range("G" & recordCount + 2).Value = "=SUM(G2:G" & recordCount & ")"
            'oDoc.Content.InsertAfter(vbCrLf & vbCrLf)
            'oDoc.Content.InsertAfter("Management Fee" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(ManagementFeePerMonth), 2) & vbCrLf)
            'invTotal += ManagementFeePerMonth
            oSheet.Range("E" & recordCount + 3).Value = "Management Fee"
            oSheet.Range("G" & recordCount + 3).Value = ManagementFeePerMonth

            'oDoc.Content.InsertAfter("Large Boxes Inbound" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(LargeBoxInbound * txtLargeInbound.Text), 2) & vbCrLf)
            'invTotal += LargeBoxInbound * txtLargeInbound.Text
            oSheet.Range("E" & recordCount + 4).Value = "Large Boxes Inbound"
            oSheet.Range("G" & recordCount + 4).Value = LargeBoxInbound * txtLargeInbound.Text

            'oDoc.Content.InsertAfter("Large Box Storage" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(LargeBoxStoragePerWeek * txtLargeStorage.Text), 2) & vbCrLf)
            'invTotal += LargeBoxStoragePerWeek * txtLargeStorage.Text
            oSheet.Range("E" & recordCount + 5).Value = "Large Box Storage"
            oSheet.Range("G" & recordCount + 5).Value = LargeBoxStoragePerWeek * txtLargeStorage.Text

            'oDoc.Content.InsertAfter("Labour" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(txtLabour.Text * HourlyRate), 2) & vbCrLf)
            'invTotal += txtLabour.Text * HourlyRate
            oSheet.Range("E" & recordCount + 6).Value = "Labour"
            oSheet.Range("G" & recordCount + 6).Value = HourlyRate * txtLabour.Text

            'oDoc.Content.InsertAfter(vbCrLf)

            'oDoc.Content.InsertAfter(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "INVOICE TOTAL" & vbTab & "£" & FormatNumber(CDbl(invTotal), 2))
            oSheet.Range("E" & recordCount + 7).Value = "Grand Total"
            oSheet.Range("G" & recordCount + 7).Value = "=SUM(G" & (recordCount + 2) & ":G" & (recordCount + 6) & ")"

            oSheet.Range("G" & (recordCount + 1) & ":G" & (recordCount + 7)).NumberFormat = "£###,###,##0.00"

            ToolStripStatusLabel1.Text = "Saving invoice..."
            oBook.SaveAs(ReportLocation & invDate.ToString("MMM yyyy") & ".xlsx")
            'oDoc.Visible = False
            oExcel.Quit()

            con.Close()
        End Using

        ToolStripStatusLabel1.Text = "Finished"

        Process.Start("explorer.exe", ReportLocation)

    End Sub
    'Private Function CalculateSpareDays(ByVal _DateIn As Date, ByVal _DateOut As Date) As Integer

    '    Dim spare As Integer = 0

    '    If invDate >= _DateOut And _DateOut <> Nothing Then
    '        For i As Integer = 1 To ((Math.Abs(_DateOut.Year - _DateIn.Year) * 12) + Math.Abs(_DateOut.Month - _DateIn.Month))
    '            If i = 1 Then
    '                spare = (Date.DaysInMonth(_DateIn.Year, _DateIn.Month) - _DateIn.Day + 1) Mod 7
    '            ElseIf i = ((Math.Abs(_DateOut.Year - _DateIn.Year) * 12) + Math.Abs(_DateOut.Month - _DateIn.Month)) Then
    '                spare = (spare + _DateOut.Day + 1) Mod 7
    '            Else
    '                Dim currentMonth = (_DateIn.Month + i - 1) Mod 12 + 1
    '                spare = (spare + Date.DaysInMonth(_DateIn.Year, currentMonth)) Mod 7
    '            End If
    '        Next
    '    ElseIf invDate < _DateOut Or _DateOut = Nothing Then
    '        For i As Integer = 1 To ((Math.Abs(invDate.Year - _DateIn.Year) * 12) + Math.Abs(invDate.Month - _DateIn.Month))
    '            If i = 1 Then
    '                spare = (Date.DaysInMonth(_DateIn.Year, _DateIn.Month) - _DateIn.Day + 1) Mod 7
    '            ElseIf i = ((Math.Abs(invDate.Year - _DateIn.Year) * 12) + Math.Abs(invDate.Month - _DateIn.Month)) Then
    '                spare = (spare + invDate.Day + 1) Mod 7
    '            Else
    '                Dim currentMonth = (_DateIn.Month + i - 1) Mod 12 + 1
    '                spare = (spare + Date.DaysInMonth(_DateIn.Year, currentMonth)) Mod 7
    '            End If
    '        Next
    '    End If
    '    Return spare
    '    'For i = 1 To ((Math.Abs(invDate.Year - )))

    'End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'invDate = DateSerial(Year(Today), Month(Today), 0)

        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "MMMMyyyy"
        DateTimePicker1.Value = DateSerial(Year(Today), Month(Today), 0)

    End Sub

    Private Sub btnDBBrowse_Click(sender As Object, e As EventArgs) Handles btnDBBrowse.Click

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Find database location"
        fd.InitialDirectory = "C:\"
        fd.Filter = "Microsoft Access (*.accdb)|*.accdb"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            txtDB.Text = strFileName
        End If

    End Sub

    Private Sub btnReportBrowse_Click(sender As Object, e As EventArgs) Handles btnReportBrowse.Click

        Dim fd As FolderBrowserDialog = New FolderBrowserDialog()
        Dim strFileName As String

        fd.Description = "Set report destination"

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.SelectedPath
            txtReports.Text = strFileName
        End If

    End Sub

    Private Sub btnSpareDaysReset_Click(sender As Object, e As EventArgs) Handles btnSpareDaysReset.Click

        ToolStripStatusLabel1.Text = "Processing spare days calculations..."
        DBLocation = txtDB.Text
        Dim SpareDate As Date = DateTimePicker1.Value
        'SpareDate = SpareDate.AddDays(Date.DaysInMonth(SpareDate.Year, SpareDate.Month) - SpareDate.Day)
        'SpareDate = SpareDate.AddMonths(-1)
        SpareDate = DateSerial(Year(SpareDate), Month(SpareDate), 0)

        con = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBLocation & ";Persist Security Info=False;")

        Dim cmd As New OleDbCommand("SELECT * FROM Goods", con)

        Using con
            Try
                con.Open()
            Catch ex As OleDb.OleDbException
                ToolStripStatusLabel1.Text = "Error: Could not find database."
                Exit Sub
            End Try
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While reader.Read()
                    Dim SerialNumber As Integer = 0
                    Dim DateIn As Date = Nothing
                    Dim DateOut As Date = Nothing
                    Dim SpareDays = 0
                    SerialNumber = reader.GetInt32(0)
                    ToolStripStatusLabel1.Text = "Processing serial number " & SerialNumber
                    Try
                        DateIn = reader.GetDateTime(2)
                    Catch ex As InvalidCastException
                        DateIn = Nothing
                    End Try
                    Try
                        DateOut = reader.GetDateTime(3)
                    Catch ex As Exception
                        DateOut = Nothing
                    End Try

                    If DateIn <> Nothing And DateIn < SpareDate Then
                        Dim CurrentDate As Date = DateIn.AddDays(Date.DaysInMonth(DateIn.Year, DateIn.Month) - DateIn.Day)
                        If DateOut = Nothing Or DateOut > SpareDate Then
                            Do
                                If CurrentDate.Month = DateIn.Month And CurrentDate.Year = DateIn.Year Then
                                    SpareDays = (CurrentDate.Day - DateIn.Day + 1) Mod 7
                                Else
                                    SpareDays = SpareDays + Date.DaysInMonth(CurrentDate.Year, CurrentDate.Month) Mod 7
                                End If
                                CurrentDate = CurrentDate.AddMonths(1)
                            Loop While CurrentDate <= SpareDate
                        ElseIf DateOut <= SpareDate Then
                            Do
                                If CurrentDate.Month = DateIn.Month And CurrentDate.Year = DateIn.Year Then
                                    SpareDays = (CurrentDate.Day - DateIn.Day + 1) Mod 7
                                ElseIf CurrentDate.Month = DateOut.Month And CurrentDate.Year = DateOut.Year Then
                                    SpareDays = (SpareDays + DateOut.Day) Mod 7
                                Else
                                    SpareDays = (SpareDays + Date.DaysInMonth(CurrentDate.Year, CurrentDate.Month)) Mod 7
                                End If
                                CurrentDate = CurrentDate.AddMonths(1)
                            Loop While CurrentDate <= (DateOut.AddDays(Date.DaysInMonth(DateOut.Year, DateOut.Month) - DateOut.Day))
                        End If
                    Else
                        SpareDays = 0
                    End If

                    Dim update As String = "UPDATE Goods SET SpareDays = " & SpareDays & " WHERE SerialNumber = " & SerialNumber
                    Dim updateCmd As New OleDbCommand(update, con)
                    updateCmd.ExecuteNonQuery()
                Loop
            End If
        End Using
        ToolStripStatusLabel1.Text = "Finished. Spare days correct to " & SpareDate
        CurrentSpareDate = SpareDate
    End Sub
End Class
