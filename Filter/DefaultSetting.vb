Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataAdapter
Imports System.Configuration
Imports System.Data

Public Class DefaultSetting

    Private Sub DefaultSetting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckedListBox1.Items.Clear()
        Try
            Dim userTables As DataTable = Nothing
            Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
            Dim restrictions() As String = New String(3) {}
            restrictions(3) = "Table"
            connection.Open()
            ' Get list of user tables
            userTables = connection.GetSchema("Tables", restrictions)
            connection.Close()
            ' Add list of table names to listBox
            ' For i As Integer = 0 To userTables.Rows.Count - 1
            ' ComboBox1.Items.Add(userTables.Rows(i)(2).ToString())
            ' set two table which is FASSETS and FADESP


            ' Console.WriteLine(userTables.Rows(i)(2).ToString())
            ' Next


            ' Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + TextBox1.Text + "'")
            Dim sql As String
            Dim cmd As New OleDb.OleDbCommand
            Dim dt As New DataTable
            Dim da As New OleDb.OleDbDataAdapter
            Dim col As String = ""
            Dim count As Integer = 0

            connection.Open()

            sql = "SELECT * FROM FASSETS"
            cmd.Connection = connection
            cmd.CommandText = sql
            da.SelectCommand = cmd

            da.Fill(dt)

            Dim dc As DataColumn

            For Each dc In dt.Columns
                If col = "" Then
                    col = "FASSETS." & dc.ColumnName
                Else
                    col = col & ", FASSETS." & dc.ColumnName
                End If

                '  ComboBox2.Items.Add(dc.ColumnName)
                ' CheckedListBox1.Items.Add(dc.ColumnName)
                ' counter = counter + 1
            Next

            Dim sql1 As String
            Dim cmd1 As New OleDb.OleDbCommand
            Dim dt1 As New DataTable
            Dim da1 As New OleDb.OleDbDataAdapter

            sql1 = "SELECT " + col + ", FADESP.Fd_Sub_No, FADESP.Fd_Description, FALOCATION.Fl_Location_ID, FALOCATION.Fl_Description, FAUSER.Fu_Description from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_User_ID = FAUSER.Fu_User_ID)"
            cmd1.Connection = connection
            cmd1.CommandText = sql1
            da1.SelectCommand = cmd1

            da1.Fill(dt1)

            Dim dc1 As DataColumn

            For Each dc1 In dt1.Columns
                CheckedListBox1.Items.Add(dc1.ColumnName)
                ' counter = counter + 1
                count = count + 1

            Next

            ToolStripStatusLabel1.Text = "Total Columns: " & count
            connection.Close()

            existingDefault()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub existingDefault()
        Dim filename As String = AppConfigReader.defaultPath
        Dim count As Integer = 0
        Using sr As New StreamReader(filename, True)
            While sr.Peek >= 0
                Dim temp As String = sr.ReadLine

                For i = 0 To CheckedListBox1.Items.Count - 1
                    If temp = CheckedListBox1.Items(i).ToString Then
                        CheckedListBox1.SetItemChecked(i, True)
                        count = count + 1

                    End If
                Next
            End While
            ToolStripStatusLabel2.Text = "Selected Columns: " & count
        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim filename As String = AppConfigReader.defaultPath

        File.WriteAllText(filename, "")

        ' If System.IO.File.Exists(filename) = True Then
        Using writer As New StreamWriter(filename, True)
            For i = 0 To CheckedListBox1.Items.Count - 1
                If CheckedListBox1.GetItemChecked(i) Then
                    writer.WriteLine(CheckedListBox1.Items(i).ToString)
                End If
            Next
        End Using

        MsgBox("Default setting saved.")

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        ToolStripStatusLabel2.Text = "Selected Columns: " & CheckedListBox1.CheckedItems.Count
    End Sub
End Class