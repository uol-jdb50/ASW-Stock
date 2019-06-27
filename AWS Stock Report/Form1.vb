Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class Form1
    Private con As OleDbConnection
    Public DBLocation As String '= "C:\Users\AB\Desktop\AWS Stock Project\"
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DBLocation = txtDB.Text
        TemplateLocation = txtReports.Text
        ReportLocation = TemplateLocation
        Do
            Dim c As Char = Strings.Right(ReportLocation, 1)
            If c = "/" Or c = "\" Then
                Exit Do
            Else
                ReportLocation = Strings.Left(ReportLocation, ReportLocation.Length - 1)
            End If
        Loop While (True)

        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBLocation & ";Persist Security Info=False;")

        Dim prices As New OleDbCommand("SELECT * FROM Charges", con)

        Using con

            con.Open()

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

        invDate = DateTimePicker1.Value
        invDate = invDate.AddDays(Date.DaysInMonth(invDate.Year, invDate.Month) - invDate.Day)

        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBLocation & ";Persist Security Info=False;")

        Dim cmd As New OleDbCommand("SELECT * FROM Goods", con)

        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        oDoc = oWord.Documents.Add(ReportLocation & "AWS Stock Report Template.docx")

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
                        Else
                            SpareDays = CalculateSpareDays(DateIn, DateOut)
                        End If

                    Catch ex As InvalidCastException
                        SpareDays = 0
                    End Try

                    If (DateIn <> Nothing And DateIn <= invDate) And (DateOut.Month = invDate.Month Or DateOut = Nothing) Then

                        oDoc.Content.InsertAfter(SerialNumber & vbTab & vbTab)
                        'Delivery inbound
                        If DateIn.Month = invDate.Month Then
                            oDoc.Content.InsertAfter("£" & FormatNumber(CDbl(DeliveryInbound), 2) & vbTab & vbTab & vbTab)
                            Total += DeliveryInbound
                        Else
                            oDoc.Content.InsertAfter(vbTab & vbTab & vbTab)
                        End If
                        'Storage
                        If DateOut.Month = invDate.Month And DateOut.Year = invDate.Year Then

                            If DateIn.Month = DateOut.Month Then
                                oDoc.Content.InsertAfter("£" & FormatNumber(CDbl((((DateOut.Day - DateIn.Day + 1 + SpareDays) \ 7) * StoragePerWeek)), 2) & vbTab & vbTab)
                                Total += ((DateOut.Day - DateIn.Day + 1 + SpareDays) \ 7) * StoragePerWeek
                            ElseIf DateIn.Month <> DateOut.Month Then
                                oDoc.Content.InsertAfter("£" & FormatNumber(CDbl((((DateOut.Day + SpareDays) \ 7) * StoragePerWeek)), 2) & vbTab & vbTab)
                                Total += ((DateOut.Day + SpareDays) \ 7) * StoragePerWeek
                            End If
                            '((DateOut.Month > invDate.Month And DateOut.Year = invDate.Year) Or (DateOut.Month = 1 And invDate.Month = 12 And DateOut.Year = invDate.Year + 1))
                        ElseIf DateOut = Nothing Or DateOut > invDate Then

                            If DateIn.Month = invDate.Month Then
                                oDoc.Content.InsertAfter("£" & FormatNumber(CDbl((((invDate.Day - DateIn.Day + 1 + SpareDays) \ 7) * StoragePerWeek)), 2) & vbTab & vbTab)
                                SpareDays = (invDate.Day - DateIn.Day + 1 + SpareDays) Mod 7
                                Total += ((invDate.Day - DateIn.Day + 1 + SpareDays) \ 7) * StoragePerWeek
                            Else
                                oDoc.Content.InsertAfter("£" & FormatNumber(CDbl((((invDate.Day + SpareDays) \ 7) * StoragePerWeek)), 2) & vbTab & vbTab)
                                SpareDays = (invDate.Day + SpareDays) Mod 7
                                Total += ((invDate.Day + SpareDays) \ 7) * StoragePerWeek
                            End If

                        ElseIf DateIn < invDate Then

                            oDoc.Content.InsertAfter("£" & FormatNumber(CDbl((((invDate.Day + SpareDays) \ 7) * StoragePerWeek)), 2) & vbTab & vbTab)
                            SpareDays = (invDate.Day + SpareDays) Mod 7
                            Total += ((invDate.Day + SpareDays) \ 7) * StoragePerWeek

                        Else

                            oDoc.Content.InsertAfter(vbTab & vbTab)

                        End If
                        'Kit check
                        If DateIn.Month = invDate.Month Then
                            If IncomingCheck = True Then
                                oDoc.Content.InsertAfter("£" & FormatNumber(CDbl(IncomingKitCheck), 2) & vbTab & vbTab)
                                Total += IncomingKitCheck
                            Else
                                oDoc.Content.InsertAfter(vbTab & vbTab)
                            End If
                        Else
                            oDoc.Content.InsertAfter(vbTab & vbTab)
                        End If
                        'Pick pack and despatch
                        If DateOut.Month = invDate.Month Then
                            oDoc.Content.InsertAfter("£" & FormatNumber(CDbl(PickPackFeePerItem), 2) & vbTab & vbTab & "£" & FormatNumber(CDbl(DespatchofKit), 2) & vbTab & vbTab)
                            Total += PickPackFeePerItem + DespatchofKit
                            DespatchTotal += DespatchofKit
                        Else
                            oDoc.Content.InsertAfter(vbTab & vbTab & vbTab & vbTab)
                        End If

                        oDoc.Content.InsertAfter("£" & FormatNumber(CDbl(Total), 2) & vbCrLf)
                        invTotal += Total

                        'UPDATE QUERY RUNS HERE
                        If invDate.Month = DateSerial(Year(Today), Month(Today), 0).Month Then
                            Dim update As String = "UPDATE Goods SET SpareDays = " & SpareDays & " WHERE SerialNumber = " & SerialNumber & ""
                            Dim updateCmd As New OleDbCommand(update, con)

                            updateCmd.ExecuteNonQuery()
                        End If
                    End If 'End of record processing

                Loop

            End If

            oDoc.Content.InsertAfter(vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "DESPATCH TOTAL" & vbTab & "£" & FormatNumber(CDbl(DespatchTotal), 2) & vbCrLf)
            oDoc.Content.InsertAfter(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SUBTOTAL" & vbTab & vbTab & "£" & FormatNumber(CDbl(invTotal), 2))

            oDoc.Content.InsertAfter(vbCrLf & vbCrLf)
            oDoc.Content.InsertAfter("Management Fee" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(ManagementFeePerMonth), 2) & vbCrLf)
            invTotal += ManagementFeePerMonth

            oDoc.Content.InsertAfter("Large Boxes Inbound" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(LargeBoxInbound * txtLargeInbound.Text), 2) & vbCrLf)
            invTotal += LargeBoxInbound * txtLargeInbound.Text

            oDoc.Content.InsertAfter("Large Box Storage" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(LargeBoxStoragePerWeek * txtLargeStorage.Text), 2) & vbCrLf)
            invTotal += LargeBoxStoragePerWeek * txtLargeStorage.Text

            oDoc.Content.InsertAfter("Labour" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "£" & FormatNumber(CDbl(txtLabour.Text * HourlyRate), 2) & vbCrLf)
            invTotal += txtLabour.Text * HourlyRate

            oDoc.Content.InsertAfter(vbCrLf)

            oDoc.Content.InsertAfter(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "INVOICE TOTAL" & vbTab & "£" & FormatNumber(CDbl(invTotal), 2))

            ToolStripStatusLabel1.Text = "Saving invoice..."
            oDoc.SaveAs2(ReportLocation & invDate.ToString("MMM yyyy") & ".docx")
            'oDoc.Visible = False
            oDoc.Close()
            oWord.Quit()

            con.Close()
        End Using

        ToolStripStatusLabel1.Text = "Finished"

        Process.Start("explorer.exe", ReportLocation)

    End Sub
    Private Function CalculateSpareDays(ByVal _DateIn As Date, ByVal _DateOut As Date) As Integer

        Dim spare As Integer = 0

        If invDate >= _DateOut And _DateOut <> Nothing Then
            For i As Integer = 1 To ((Math.Abs(_DateOut.Year - _DateIn.Year) * 12) + Math.Abs(_DateOut.Month - _DateIn.Month))
                If i = 1 Then
                    spare = (Date.DaysInMonth(_DateIn.Year, _DateIn.Month) - _DateIn.Day + 1) Mod 7
                ElseIf i = ((Math.Abs(_DateOut.Year - _DateIn.Year) * 12) + Math.Abs(_DateOut.Month - _DateIn.Month)) Then
                    spare = (spare + _DateOut.Day + 1) Mod 7
                Else
                    Dim currentMonth = (_DateIn.Month + i - 1) Mod 12 + 1
                    spare = (spare + Date.DaysInMonth(_DateIn.Year, currentMonth)) Mod 7
                End If
            Next
        ElseIf invDate < _DateOut Or _DateOut = Nothing Then
            For i As Integer = 1 To ((Math.Abs(invDate.Year - _DateIn.Year) * 12) + Math.Abs(invDate.Month - _DateIn.Month))
                If i = 1 Then
                    spare = (Date.DaysInMonth(_DateIn.Year, _DateIn.Month) - _DateIn.Day + 1) Mod 7
                ElseIf i = ((Math.Abs(invDate.Year - _DateIn.Year) * 12) + Math.Abs(invDate.Month - _DateIn.Month)) Then
                    spare = (spare + invDate.Day + 1) Mod 7
                Else
                    Dim currentMonth = (_DateIn.Month + i - 1) Mod 12 + 1
                    spare = (spare + Date.DaysInMonth(_DateIn.Year, currentMonth)) Mod 7
                End If
            Next
        End If
        Return spare
        'For i = 1 To ((Math.Abs(invDate.Year - )))

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'invDate = DateSerial(Year(Today), Month(Today), 0)

        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "MMMMyyyy"

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

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Find reports template"
        fd.InitialDirectory = "C:\"
        fd.Filter = "Microsoft Word (*.docx)|*.docx"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            txtReports.Text = strFileName
        End If

    End Sub
End Class
