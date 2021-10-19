Imports System.Data.OleDb

Public Class Form1

    Private Ref As DBConnect.OleDBCommand =
    New DBConnect.OleDBCommand("Provider=SQLOLEDB;Data Source=localhost\MSSQL2017EXPR; Initial Catalog=ST_DB;Integrated Security=SSPI;")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim tb As DataTable = New DataTable

        Dim sql As List(Of String) = New List(Of String)
        sql.Add("UPDATE Table_1 SET float = 19910618 WHERE int = 1")
        sql.Add("delete from Table_1 WHERE int = 2")
        sql.Add("INSERT INTO Table_1 VALUES (3,'nishi','2020-01-31',1000)")

        If Ref.OleDbExcuteNonQuery(sql) = False Then
            MessageBox.Show("エラーが発生しました。")
            Exit Sub
        End If

        MessageBox.Show("正常に終了しました。")

    End Sub
End Class


'意図したテーブルへ接続できているかどうかの確認用のコード
'Dim tb As DataTable = New DataTable

'Dim sql As List(Of String) = New List(Of String)
'sql.Add("select * from Table_1")

'tb = Ref.OleDBDataTable(sql(0).ToString)

'MessageBox.Show(tb.Rows(0).Item("Datetime"))
'MessageBox.Show(tb.Rows(0).Item("Float"))


'検証①のソース（ボタンクリックイベント内）
'Dim tb As DataTable = New DataTable

'Dim sql As List(Of String) = New List(Of String)
'sql.Add("UPDATE Table_1 SET float = 19910618 WHERE int = 1")
'sql.Add("delete from Table_1 WHERE int = 2")
'sql.Add("INSERT INTO Table_1 VALUES (3,'nishi','2020-01-31',1000)")

'If Ref.OleDbExcuteNonQuery(sql) = False Then
'MessageBox.Show("エラーが発生しました。")
'Exit Sub
'End If

'MessageBox.Show("正常に終了しました。")

'検証②ソース（ボタンクリックイベント内）

'Dim tb As DataTable = New DataTable

'Dim sql As List(Of String) = New List(Of String)
'sql.Add("UPDATE Table_1 SET float = 20080618 WHERE int = 1")
'sql.Add("delete from Table_1 WHERE int = 3")
'sql.Add("INSERT INTO Table_1 VALUES (100,'nishi','2020-01-31','das')")

'If Ref.OleDbExcuteNonQuery(sql) = False Then
'MessageBox.Show("エラーが発生しました。")
'Exit Sub
'End If

'MessageBox.Show("正常に終了しました。")