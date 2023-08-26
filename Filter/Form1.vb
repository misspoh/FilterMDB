Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataAdapter
Imports System.Configuration
Imports System.Data

Public Class Form1
    Dim datat As DataTable
    Dim counter As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim userTables As DataTable = Nothing
            Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
            ' We only want the user tables, not system tables
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

            sql1 = "SELECT " + col + ", FADESP.Fd_Sub_No, FADESP.Fd_Description, FALOCATION.Fl_Location_ID, FALOCATION.Fl_Description, FAUSER.Fu_Description from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID)"
            cmd1.Connection = connection
            cmd1.CommandText = sql1
            da1.SelectCommand = cmd1

            da1.Fill(dt1)

            Dim dc1 As DataColumn

            For Each dc1 In dt1.Columns
                CheckedListBox1.Items.Add(dc1.ColumnName)
                ' counter = counter + 1
            Next

            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
       
    End Sub

    Sub countChecked()
        
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            DataGridView1.DataSource = Nothing
            ' ComboBox2.Items.Clear()
            TextBox2.Text = ""
            TextBox3.Text = ""

            Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
            Dim sql As String
            Dim cmd As New OleDb.OleDbCommand
            Dim dt As New DataTable
            Dim da As New OleDb.OleDbDataAdapter
            Dim col As String = ""

            For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                If (CheckedListBox1.GetItemChecked(j)) Then
                    If col = "" Then
                        col = "FASSETS." & CheckedListBox1.Items(j).ToString
                    Else
                        Dim temp As String = CheckedListBox1.Items(j).ToString

                        If temp.Substring(0, 2) = "Fd" Then
                            col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                        ElseIf temp.Substring(0, 2) = "Fl" Then
                            col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                        ElseIf temp.Substring(0, 2) = "Fu" Then
                            col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                        Else

                            col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                        End If

                        If temp.Length > 6 Then
                            If temp.Substring(3, 4) = "DATE" Then
                                col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                            End If
                        End If

                    End If
                End If
            Next

            ' MsgBox(col)

            connection.Open()
            sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID)"
            cmd.Connection = connection
            cmd.CommandText = sql
            da.SelectCommand = cmd

            da.Fill(dt)

            DataGridView1.DataSource = dt
            dgvatm()
            ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

            rightalign()
            'Dim dc As DataColumn
            'Dim ds As New DataSet

            ''da.Fill(ds, ComboBox1.Text)

            '' CheckedListBox1.Items.Add("Select All")

            'For Each dc In dt.Columns

            '    ' ComboBox2.Items.Add(dc.ColumnName)
            '    '  CheckedListBox1.Items.Add(dc.ColumnName)
            '    counter = counter + 1
            'Next

            CheckBox2.Enabled = True
            CheckBox4.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            Button4.Enabled = True

            connection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
       
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If CheckBox2.CheckState = CheckState.Checked Then
            If CheckBox4.CheckState = CheckState.Checked Then
                DataGridView1.DataSource = Nothing

                Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                Dim sql As String = ""
                Dim cmd As New OleDb.OleDbCommand
                Dim dt As New DataTable
                Dim da As New OleDb.OleDbDataAdapter
                Dim col As String = ""

                For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                    If (CheckedListBox1.GetItemChecked(j)) Then
                        If col = "" Then
                            col = "FASSETS." & CheckedListBox1.Items(j).ToString
                        Else
                            Dim temp As String = CheckedListBox1.Items(j).ToString

                            If temp.Substring(0, 2) = "Fd" Then
                                col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fl" Then
                                col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fu" Then
                                col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                            Else

                                col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                            End If

                            If temp.Length > 6 Then
                                If temp.Substring(3, 4) = "DATE" Then
                                    col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                End If

                            End If
                        End If
                    End If
                Next

                'MsgBox(col)

                connection.Open()

                If TextBox2.Text = "" And TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"
                ElseIf TextBox2.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"
                ElseIf TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"
                Else
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"

                End If

                cmd.Connection = connection
                cmd.CommandText = sql
                da.SelectCommand = cmd

                da.Fill(dt)

                DataGridView1.DataSource = dt
                dgvatm()
                ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                connection.Close()
                'connection.Open()

                'sql = "SELECT * FROM " + ComboBox1.Text + " WHERE " ' + ComboBox2.Text + " LIKE '" + TextBox2.Text + "%'"
                'cmd.Connection = connection
                'cmd.CommandText = sql
                'da.SelectCommand = cmd

                'da.Fill(dt)

                'DataGridView1.DataSource = dt

                'datat = dt
            Else
                DataGridView1.DataSource = Nothing

                Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                Dim sql As String = ""
                Dim cmd As New OleDb.OleDbCommand
                Dim dt As New DataTable
                Dim da As New OleDb.OleDbDataAdapter
                Dim col As String = ""

                For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                    If (CheckedListBox1.GetItemChecked(j)) Then
                        If col = "" Then
                            col = "FASSETS." & CheckedListBox1.Items(j).ToString
                        Else
                            Dim temp As String = CheckedListBox1.Items(j).ToString

                            If temp.Substring(0, 2) = "Fd" Then
                                col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fl" Then
                                col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fu" Then
                                col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                            Else

                                col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                            End If

                            If temp.Length > 6 Then
                                If temp.Substring(3, 4) = "DATE" Then
                                    col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                End If

                            End If
                        End If
                    End If
                Next

                'MsgBox(col)

                connection.Open()

                If TextBox2.Text = "" And TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FADESP.Fd_Sub_No = 0"
                ElseIf TextBox2.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0"
                ElseIf TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FADESP.Fd_Sub_No = 0"
                Else
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0"

                End If

                cmd.Connection = connection
                cmd.CommandText = sql
                da.SelectCommand = cmd

                da.Fill(dt)

                DataGridView1.DataSource = dt
                dgvatm()
                ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                connection.Close()
                'connection.Open()

                'sql = "SELECT * FROM " + ComboBox1.Text + " WHERE " ' + ComboBox2.Text + " LIKE '" + TextBox2.Text + "%'"
                'cmd.Connection = connection
                'cmd.CommandText = sql
                'da.SelectCommand = cmd

                'da.Fill(dt)

                'DataGridView1.DataSource = dt

                'datat = dt
            End If
           
        Else
            If CheckBox4.CheckState = CheckState.Checked Then
                DataGridView1.DataSource = Nothing

                Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                Dim sql As String = ""
                Dim cmd As New OleDb.OleDbCommand
                Dim dt As New DataTable
                Dim da As New OleDb.OleDbDataAdapter
                Dim col As String = ""

                For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                    If (CheckedListBox1.GetItemChecked(j)) Then
                        If col = "" Then
                            col = "FASSETS." & CheckedListBox1.Items(j).ToString
                        Else
                            Dim temp As String = CheckedListBox1.Items(j).ToString

                            If temp.Substring(0, 2) = "Fd" Then
                                col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fl" Then
                                col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fu" Then
                                col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                            Else

                                col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                            End If

                            If temp.Length > 6 Then
                                If temp.Substring(3, 4) = "DATE" Then
                                    col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                End If

                            End If
                        End If
                    End If

                Next

                'MsgBox(col)

                connection.Open()

                If TextBox2.Text = "" And TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Quantity > 0"
                ElseIf TextBox2.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0"
                ElseIf TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Quantity > 0"
                Else
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0"
                End If

                cmd.Connection = connection
                cmd.CommandText = sql
                da.SelectCommand = cmd

                da.Fill(dt)

                DataGridView1.DataSource = dt
                dgvatm()
                ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                connection.Close()

                'connection.Open()

                'sql = "SELECT * FROM " + ComboBox1.Text + " WHERE " ' + ComboBox2.Text + " LIKE '" + TextBox2.Text + "%'"
                'cmd.Connection = connection
                'cmd.CommandText = sql
                'da.SelectCommand = cmd

                'da.Fill(dt)

                'DataGridView1.DataSource = dt

                'datat = dt

            Else
                DataGridView1.DataSource = Nothing

                Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                Dim sql As String = ""
                Dim cmd As New OleDb.OleDbCommand
                Dim dt As New DataTable
                Dim da As New OleDb.OleDbDataAdapter
                Dim col As String = ""

                For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                    If (CheckedListBox1.GetItemChecked(j)) Then
                        If col = "" Then
                            col = "FASSETS." & CheckedListBox1.Items(j).ToString
                        Else
                            Dim temp As String = CheckedListBox1.Items(j).ToString

                            If temp.Substring(0, 2) = "Fd" Then
                                col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fl" Then
                                col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                            ElseIf temp.Substring(0, 2) = "Fu" Then
                                col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                            Else

                                col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                            End If

                            If temp.Length > 6 Then
                                If temp.Substring(3, 4) = "DATE" Then
                                    col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                End If

                            End If
                        End If
                    End If

                Next

                'MsgBox(col)

                connection.Open()

                If TextBox2.Text = "" And TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID)"
                ElseIf TextBox2.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%'"
                ElseIf TextBox3.Text = "" Then
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%'"
                Else
                    sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON Fassets.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%'"
                End If

                cmd.Connection = connection
                cmd.CommandText = sql
                da.SelectCommand = cmd

                da.Fill(dt)

                DataGridView1.DataSource = dt
                dgvatm()
                ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                connection.Close()

                'connection.Open()

                'sql = "SELECT * FROM " + ComboBox1.Text + " WHERE " ' + ComboBox2.Text + " LIKE '" + TextBox2.Text + "%'"
                'cmd.Connection = connection
                'cmd.CommandText = sql
                'da.SelectCommand = cmd

                'da.Fill(dt)

                'DataGridView1.DataSource = dt

                'datat = dt

            End If
           
        End If
    End Sub

    Sub dgvatm()
        For i As Integer = 0 To DataGridView1.ColumnCount - 1
            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        Next
    End Sub

    Private Sub btn_Reset_Click(sender As Object, e As EventArgs) Handles btn_Reset.Click
        DataGridView1.DataSource = Nothing

        'TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

        ' ComboBox1.Items.Clear()
        'ComboBox2.Items.Clear()

        ' ComboBox1.Text = ""
        'ComboBox2.Text = ""

        ' CheckedListBox1.Items.Clear()

        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, False)
        Next

        CheckBox2.CheckState = CheckState.Unchecked
        ToolStripStatusLabel1.Text = "Total:"

        CheckBox3.CheckState = CheckState.Unchecked

        CheckBox2.Enabled = False
        CheckBox4.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        Button4.Enabled = False

        counter = 0
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        SaveFileDialog1.Filter = "Excel Workbook|*.xlsx"

        ProgressBar1.Maximum = DataGridView1.RowCount

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Try

            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "Message")
            'End Try
            Dim xlApp As Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
            Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim i As Integer
            Dim j As Integer

            xlApp = New Microsoft.Office.Interop.Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")

            For i = 0 To DataGridView1.RowCount - 1
                For j = 0 To DataGridView1.ColumnCount - 1

                    Dim temp As String = DataGridView1.Columns(j).HeaderText

                    xlWorkSheet.Cells(1, j + 1) = DataGridView1.Columns(j).HeaderText

                    If temp.Contains("Date") Then
                        xlWorkSheet.Cells(i + 2, j + 1) = Format(CDate(DataGridView1(j, i).Value.ToString()), "dd-MM-yyyy")

                    Else
                        xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
                    End If

                    '    For k As Integer = 1 To DataGridView1.Columns.Count
                    '        ' date issues

                    '        xlWorkSheet.Cells(1, j) = DataGridView1.Columns(j).HeaderText
                    '        xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()

                    '        ' If temp.Substring (3, 4) = 
                    '    Next

                Next
                ProgressBar1.Value = i

            Next

            Dim fullpath As String = SaveFileDialog1.FileName
            xlWorkSheet.SaveAs(fullpath)
            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

            MsgBox("Successfully export.")
        End If

        ProgressBar1.Value = 0

        'releaseObject(xlApp)
        'releaseObject(xlWorkBook)
        'releaseObject(xlWorkSheet)
    End Sub

    Sub exportToExcel()
        Dim dt As DataTable = DataGridView1.DataSource

        Dim dataset As DataSet = New DataSet
        dataset.Tables.Add(dt)

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, True)
            Next
        Else
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If
    End Sub

    Private Sub CheckBox2_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckStateChanged
        

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles btnDefault.Click
        DefaultSetting.ShowDialog()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.CheckState = CheckState.Checked Then
            Dim filename As String = AppConfigReader.defaultPath

            Using sr As New StreamReader(filename, True)
                While sr.Peek >= 0
                    Dim temp As String = sr.ReadLine

                    For i = 0 To CheckedListBox1.Items.Count - 1
                        If temp = CheckedListBox1.Items(i).ToString Then
                            CheckedListBox1.SetItemChecked(i, True)
                        End If
                    Next
                End While
            End Using
            btnDefault.Enabled = False
        Else
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
            btnDefault.Enabled = True
        End If

    End Sub

    Sub dtToExcel()

    End Sub

    Private Sub CheckBox4_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckStateChanged
      

    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        If CheckedListBox1.Items.Count = Nothing Then
        Else
            If CheckBox2.CheckState = CheckState.Checked Then
                If CheckBox4.CheckState = CheckState.Checked Then
                    For i As Integer = 0 To DataGridView1.Columns.Count - 1
                        If DataGridView1.Columns(i).HeaderText = "Fd_Sub_No" Then
                            DataGridView1.DataSource = Nothing

                            Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                            Dim sql As String = ""
                            Dim cmd As New OleDb.OleDbCommand
                            Dim dt As New DataTable
                            Dim da As New OleDb.OleDbDataAdapter
                            Dim col As String = ""

                            For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                                If (CheckedListBox1.GetItemChecked(j)) Then
                                    If col = "" Then
                                        col = "FASSETS." & CheckedListBox1.Items(j).ToString
                                    Else
                                        Dim temp As String = CheckedListBox1.Items(j).ToString

                                        If temp.Substring(0, 2) = "Fd" Then
                                            col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fl" Then
                                            col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fu" Then
                                            col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                        Else

                                            col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                        End If


                                        If temp.Length > 6 Then

                                            If temp.Substring(3, 4) = "DATE" Then
                                                col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                            End If
                                        End If
                                    End If
                                End If
                            Next

                            ' MsgBox(col)

                            connection.Open()

                            If TextBox2.Text = "" And TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"
                            ElseIf TextBox2.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"
                            ElseIf TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"
                            Else
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0 AND FASSETS.Fa_Quantity > 0"

                            End If

                            cmd.Connection = connection
                            cmd.CommandText = sql
                            da.SelectCommand = cmd

                            da.Fill(dt)

                            If dt.Rows.Count = Nothing Then

                            End If
                            DataGridView1.DataSource = dt
                            dgvatm()
                            ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                            'Dim dc As DataColumn
                            'Dim ds As New DataSet

                            ''da.Fill(ds, ComboBox1.Text)

                            '' CheckedListBox1.Items.Add("Select All")

                            'For Each dc In dt.Columns

                            '    ' ComboBox2.Items.Add(dc.ColumnName)
                            '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                            '    counter = counter + 1
                            'Next

                            connection.Close()
                        End If
                    Next
                Else
                    For i As Integer = 0 To DataGridView1.Columns.Count - 1
                        If DataGridView1.Columns(i).HeaderText = "Fd_Sub_No" Then
                            DataGridView1.DataSource = Nothing

                            Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                            Dim sql As String = ""
                            Dim cmd As New OleDb.OleDbCommand
                            Dim dt As New DataTable
                            Dim da As New OleDb.OleDbDataAdapter
                            Dim col As String = ""

                            For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                                If (CheckedListBox1.GetItemChecked(j)) Then
                                    If col = "" Then
                                        col = "FASSETS." & CheckedListBox1.Items(j).ToString
                                    Else
                                        Dim temp As String = CheckedListBox1.Items(j).ToString

                                        If temp.Substring(0, 2) = "Fd" Then
                                            col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fl" Then
                                            col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fu" Then
                                            col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                        Else

                                            col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                        End If


                                        If temp.Length > 6 Then

                                            If temp.Substring(3, 4) = "DATE" Then
                                                col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                            End If
                                        End If
                                    End If
                                End If
                            Next

                            ' MsgBox(col)

                            connection.Open()

                            If TextBox2.Text = "" And TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FADESP.Fd_Sub_No = 0 "
                            ElseIf TextBox2.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0"
                            ElseIf TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FADESP.Fd_Sub_No = 0"
                            Else
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0"

                            End If

                            cmd.Connection = connection
                            cmd.CommandText = sql
                            da.SelectCommand = cmd

                            da.Fill(dt)

                            If dt.Rows.Count = Nothing Then

                            End If
                            DataGridView1.DataSource = dt
                            dgvatm()
                            ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                            'Dim dc As DataColumn
                            'Dim ds As New DataSet

                            ''da.Fill(ds, ComboBox1.Text)

                            '' CheckedListBox1.Items.Add("Select All")

                            'For Each dc In dt.Columns

                            '    ' ComboBox2.Items.Add(dc.ColumnName)
                            '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                            '    counter = counter + 1
                            'Next

                            connection.Close()
                        End If
                    Next
                End If

            Else
                If CheckBox4.CheckState = CheckState.Checked Then
                    DataGridView1.DataSource = Nothing

                    Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                    Dim sql As String = ""
                    Dim cmd As New OleDb.OleDbCommand
                    Dim dt As New DataTable
                    Dim da As New OleDb.OleDbDataAdapter
                    Dim col As String = ""

                    For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                        If (CheckedListBox1.GetItemChecked(j)) Then
                            If col = "" Then
                                col = "FASSETS." & CheckedListBox1.Items(j).ToString
                            Else
                                Dim temp As String = CheckedListBox1.Items(j).ToString

                                If temp.Substring(0, 2) = "Fd" Then
                                    col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fl" Then
                                    col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fu" Then
                                    col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                Else

                                    col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                End If


                                If temp.Length > 6 Then
                                    If temp.Substring(3, 4) = "DATE" Then
                                        col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                    End If

                                End If
                            End If
                        End If
                    Next

                    ' MsgBox(col)

                    connection.Open()

                    If TextBox2.Text = "" And TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Quantity > 0 "
                    ElseIf TextBox2.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0"
                    ElseIf TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Quantity > 0"
                    Else
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0"

                    End If

                    cmd.Connection = connection
                    cmd.CommandText = sql
                    da.SelectCommand = cmd

                    da.Fill(dt)

                    DataGridView1.DataSource = dt
                    dgvatm()
                    ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                    'Dim dc As DataColumn
                    'Dim ds As New DataSet

                    ''da.Fill(ds, ComboBox1.Text)

                    '' CheckedListBox1.Items.Add("Select All")

                    'For Each dc In dt.Columns

                    '    ' ComboBox2.Items.Add(dc.ColumnName)
                    '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                    '    counter = counter + 1
                    'Next

                    connection.Close()
                Else
                    DataGridView1.DataSource = Nothing

                    Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                    Dim sql As String = ""
                    Dim cmd As New OleDb.OleDbCommand
                    Dim dt As New DataTable
                    Dim da As New OleDb.OleDbDataAdapter
                    Dim col As String = ""

                    For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                        If (CheckedListBox1.GetItemChecked(j)) Then
                            If col = "" Then
                                col = "FASSETS." & CheckedListBox1.Items(j).ToString
                            Else
                                Dim temp As String = CheckedListBox1.Items(j).ToString

                                If temp.Substring(0, 2) = "Fd" Then
                                    col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fl" Then
                                    col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fu" Then
                                    col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                Else

                                    col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                End If


                                If temp.Length > 6 Then
                                    If temp.Substring(3, 4) = "DATE" Then
                                        col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                    End If

                                End If
                            End If
                        End If
                    Next

                    ' MsgBox(col)

                    connection.Open()

                    If TextBox2.Text = "" And TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) "
                    ElseIf TextBox2.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%'"
                    ElseIf TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%'"
                    Else
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%'"

                    End If

                    cmd.Connection = connection
                    cmd.CommandText = sql
                    da.SelectCommand = cmd

                    da.Fill(dt)

                    DataGridView1.DataSource = dt
                    dgvatm()
                    ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                    'Dim dc As DataColumn
                    'Dim ds As New DataSet

                    ''da.Fill(ds, ComboBox1.Text)

                    '' CheckedListBox1.Items.Add("Select All")

                    'For Each dc In dt.Columns

                    '    ' ComboBox2.Items.Add(dc.ColumnName)
                    '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                    '    counter = counter + 1
                    'Next

                    connection.Close()
                End If
              
            End If
            rightalign()
        End If
    End Sub

    Private Sub CheckBox4_Click(sender As Object, e As EventArgs) Handles CheckBox4.Click
        If CheckedListBox1.Items.Count = Nothing Then
        Else
            If CheckBox4.CheckState = CheckState.Checked Then
                If CheckBox2.CheckState = CheckState.Checked Then
                    For i As Integer = 0 To DataGridView1.Columns.Count - 1
                        If DataGridView1.Columns(i).HeaderText = "Fa_Quantity" Then
                            DataGridView1.DataSource = Nothing

                            Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                            Dim sql As String = ""
                            Dim cmd As New OleDb.OleDbCommand
                            Dim dt As New DataTable
                            Dim da As New OleDb.OleDbDataAdapter
                            Dim col As String = ""

                            For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                                If (CheckedListBox1.GetItemChecked(j)) Then
                                    If col = "" Then
                                        col = "FASSETS." & CheckedListBox1.Items(j).ToString
                                    Else
                                        Dim temp As String = CheckedListBox1.Items(j).ToString

                                        If temp.Substring(0, 2) = "Fd" Then
                                            col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fl" Then
                                            col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fu" Then
                                            col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                        Else

                                            col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                        End If

                                        If temp.Length > 6 Then
                                            If temp.Substring(3, 4) = "DATE" Then
                                                col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                            End If


                                        End If
                                    End If
                                End If
                            Next

                            ' MsgBox(col)

                            connection.Open()

                            If TextBox2.Text = "" And TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Quantity > 0 AND FADESP.Fd_Sub_No = 0"
                            ElseIf TextBox2.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0 AND FADESP.Fd_Sub_No = 0"
                            ElseIf TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Quantity > 0 AND FADESP.Fd_Sub_No = 0"
                            Else
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0 AND FADESP.Fd_Sub_No = 0"

                            End If

                            cmd.Connection = connection
                            cmd.CommandText = sql
                            da.SelectCommand = cmd

                            da.Fill(dt)

                            If dt.Rows.Count = Nothing Then

                            End If
                            DataGridView1.DataSource = dt
                            dgvatm()
                            ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                            'Dim dc As DataColumn
                            'Dim ds As New DataSet

                            ''da.Fill(ds, ComboBox1.Text)

                            '' CheckedListBox1.Items.Add("Select All")

                            'For Each dc In dt.Columns

                            '    ' ComboBox2.Items.Add(dc.ColumnName)
                            '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                            '    counter = counter + 1
                            'Next

                            connection.Close()
                        End If
                    Next
                Else
                    For i As Integer = 0 To DataGridView1.Columns.Count - 1
                        If DataGridView1.Columns(i).HeaderText = "Fa_Quantity" Then
                            DataGridView1.DataSource = Nothing

                            Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                            Dim sql As String = ""
                            Dim cmd As New OleDb.OleDbCommand
                            Dim dt As New DataTable
                            Dim da As New OleDb.OleDbDataAdapter
                            Dim col As String = ""

                            For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                                If (CheckedListBox1.GetItemChecked(j)) Then
                                    If col = "" Then
                                        col = "FASSETS." & CheckedListBox1.Items(j).ToString
                                    Else
                                        Dim temp As String = CheckedListBox1.Items(j).ToString

                                        If temp.Substring(0, 2) = "Fd" Then
                                            col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fl" Then
                                            col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                        ElseIf temp.Substring(0, 2) = "Fu" Then
                                            col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                        Else

                                            col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                        End If

                                        If temp.Length > 6 Then
                                            If temp.Substring(3, 4) = "DATE" Then
                                                col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                            End If


                                        End If
                                    End If
                                End If
                            Next

                            ' MsgBox(col)

                            connection.Open()

                            If TextBox2.Text = "" And TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Quantity > 0"
                            ElseIf TextBox2.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0"
                            ElseIf TextBox3.Text = "" Then
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Quantity > 0"
                            Else
                                sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FASSETS.Fa_Quantity > 0"

                            End If

                            cmd.Connection = connection
                            cmd.CommandText = sql
                            da.SelectCommand = cmd

                            da.Fill(dt)

                            If dt.Rows.Count = Nothing Then

                            End If
                            DataGridView1.DataSource = dt
                            dgvatm()
                            ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                            'Dim dc As DataColumn
                            'Dim ds As New DataSet

                            ''da.Fill(ds, ComboBox1.Text)

                            '' CheckedListBox1.Items.Add("Select All")

                            'For Each dc In dt.Columns

                            '    ' ComboBox2.Items.Add(dc.ColumnName)
                            '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                            '    counter = counter + 1
                            'Next

                            connection.Close()
                        End If
                    Next
                End If

            Else
                If CheckBox2.CheckState = CheckState.Checked Then
                    DataGridView1.DataSource = Nothing

                    Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                    Dim sql As String = ""
                    Dim cmd As New OleDb.OleDbCommand
                    Dim dt As New DataTable
                    Dim da As New OleDb.OleDbDataAdapter
                    Dim col As String = ""

                    For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                        If (CheckedListBox1.GetItemChecked(j)) Then
                            If col = "" Then
                                col = "FASSETS." & CheckedListBox1.Items(j).ToString
                            Else
                                Dim temp As String = CheckedListBox1.Items(j).ToString

                                If temp.Substring(0, 2) = "Fd" Then
                                    col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fl" Then
                                    col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fu" Then
                                    col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                Else

                                    col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                End If

                                If temp.Length > 6 Then
                                    If temp.Substring(3, 4) = "DATE" Then
                                        col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                    End If

                                End If
                            End If
                        End If
                    Next

                    ' MsgBox(col)

                    connection.Open()

                    If TextBox2.Text = "" And TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FADESP.Fd_Sub_No = 0"
                    ElseIf TextBox2.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0"
                    ElseIf TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FADESP.Fd_Sub_No = 0"
                    Else
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%' AND FADESP.Fd_Sub_No = 0"

                    End If

                    cmd.Connection = connection
                    cmd.CommandText = sql
                    da.SelectCommand = cmd

                    da.Fill(dt)

                    DataGridView1.DataSource = dt
                    dgvatm()
                    ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                    'Dim dc As DataColumn
                    'Dim ds As New DataSet

                    ''da.Fill(ds, ComboBox1.Text)

                    '' CheckedListBox1.Items.Add("Select All")

                    'For Each dc In dt.Columns

                    '    ' ComboBox2.Items.Add(dc.ColumnName)
                    '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                    '    counter = counter + 1
                    'Next

                    connection.Close()
                Else
                    DataGridView1.DataSource = Nothing

                    Dim connection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + AppConfigReader.filePath + "'")
                    Dim sql As String = ""
                    Dim cmd As New OleDb.OleDbCommand
                    Dim dt As New DataTable
                    Dim da As New OleDb.OleDbDataAdapter
                    Dim col As String = ""

                    For j As Integer = 0 To CheckedListBox1.Items.Count - 1
                        If (CheckedListBox1.GetItemChecked(j)) Then
                            If col = "" Then
                                col = "FASSETS." & CheckedListBox1.Items(j).ToString
                            Else
                                Dim temp As String = CheckedListBox1.Items(j).ToString

                                If temp.Substring(0, 2) = "Fd" Then
                                    col = col & ", FADESP." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fl" Then
                                    col = col & ", FALOCATION." & CheckedListBox1.Items(j).ToString
                                ElseIf temp.Substring(0, 2) = "Fu" Then
                                    col = col & ", FAUSER." & CheckedListBox1.Items(j).ToString
                                Else

                                    col = col & ", FASSETS." & CheckedListBox1.Items(j).ToString
                                End If

                                If temp.Length > 6 Then
                                    If temp.Substring(3, 4) = "DATE" Then
                                        col = ", CONVERT(VARCHAR(10), FASSETS." & CheckedListBox1.Items(j).ToString & ", 104) as " & CheckedListBox1.Items(j).ToString
                                    End If

                                End If
                            End If
                        End If
                    Next

                    ' MsgBox(col)

                    connection.Open()

                    If TextBox2.Text = "" And TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID)"
                    ElseIf TextBox2.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%'"
                    ElseIf TextBox3.Text = "" Then
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%'"
                    Else
                        sql = "SELECT " + col + " from ((( [FASSETS] INNER JOIN FADESP ON FASSETS.Fa_Asset_ID = FADESP.Fd_Asset_ID) INNER JOIN FALOCATION ON FASSETS.Fa_Location_ID = FALOCATION.Fl_Location_ID) INNER JOIN FAUSER ON FASSETS.Fa_Asset_User_ID = FAUSER.Fu_User_ID) WHERE FASSETS.Fa_Location_ID LIKE '" + TextBox2.Text + "%' AND FASSETS.Fa_Class_ID LIKE '" + TextBox3.Text + "%'"

                    End If

                    cmd.Connection = connection
                    cmd.CommandText = sql
                    da.SelectCommand = cmd

                    da.Fill(dt)

                    DataGridView1.DataSource = dt
                    dgvatm()
                    ToolStripStatusLabel1.Text = "Total Rows: " & dt.Rows.Count

                    'Dim dc As DataColumn
                    'Dim ds As New DataSet

                    ''da.Fill(ds, ComboBox1.Text)

                    '' CheckedListBox1.Items.Add("Select All")

                    'For Each dc In dt.Columns

                    '    ' ComboBox2.Items.Add(dc.ColumnName)
                    '    '  CheckedListBox1.Items.Add(dc.ColumnName)
                    '    counter = counter + 1
                    'Next

                    connection.Close()
                End If
             
            End If
            rightalign()
        End If
    End Sub

    Sub rightalign()
        For i As Integer = 0 To DataGridView1.Columns.Count - 1
            If IsDBNull(DataGridView1.Item(i, 0).Value) Then
            Else
                Dim temp As String = DataGridView1.Item(i, 0).Value

                If IsNumeric(temp) Then
                    DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End If
            End If

        Next
    End Sub

    
End Class
