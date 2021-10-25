Imports System.Windows.Forms

Public Class OleDBCommand

    Private _sqlConnect As String '接続文字を格納します。

    ''' <summary>
    ''' 接続文字を指定せずインスタンス化を行います。
    ''' </summary>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' 接続文字を指定してインスタンス化を行います。
    ''' </summary>
    ''' <param name="sqlConnect">接続文字</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sqlConnect As String)
        _sqlConnect = sqlConnect
    End Sub

    ''' <summary>
    ''' 接続文字を格納します。
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property SqlConnect()
        Set(ByVal value)
            _sqlConnect = value
        End Set
    End Property

    ''' <summary>
    ''' SQLコマンドを実行し、実行結果を取得します。
    ''' </summary>
    ''' <param name="sqlCommand">実行するSQL文（SELECT文)</param>
    ''' <returns>実行結果</returns>
    ''' <remarks></remarks>
    Public Function OleDBDataTable(sqlCommand As String) As DataTable

        Dim cn As OleDb.OleDbConnection = New OleDb.OleDbConnection(_sqlConnect)
        Dim Com As OleDb.OleDbCommand = New OleDb.OleDbCommand(sqlCommand, cn)
        Dim respTable As DataTable = New DataTable

        cn.Open()

        Try
            respTable.Load(Com.ExecuteReader)

            Return respTable

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        Finally
            cn.Close()
        End Try

    End Function

    ''' <summary>
    ''' SQLコマンドを実行します。
    ''' </summary>
    ''' <param name="sqlCommands">実行するSQL文(UPDATE又はDELETE文)</param>
    ''' <returns>実行件数</returns>
    ''' <remarks></remarks>
    Public Function OleDbExcuteNonQuery(ParamArray sqlCommands As String()) As Boolean

        Dim cn As OleDb.OleDbConnection = New OleDb.OleDbConnection(_sqlConnect)
        Dim OleTran As OleDb.OleDbTransaction
        Dim Com As OleDb.OleDbCommand = New OleDb.OleDbCommand()
        Com.Connection = cn

        cn.Open()
        OleTran = cn.BeginTransaction
        Com.Transaction = OleTran

        Try

            Dim i1 As Long
            For i1 = 0 To sqlCommands.Count - 1
                Com.CommandText = sqlCommands(i1)
                Com.ExecuteNonQuery()
            Next

            OleTran.Commit()
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OleTran.Rollback()
            Return False
        Finally
            OleTran.Dispose()
            cn.Close()
        End Try

    End Function

    ''' <summary>
    ''' SQLコマンドを実行し、実行結果を1つだけ取得します。
    ''' </summary>
    ''' <param name="sqlCommand">実行結果が1行1列となるSQL文</param>
    ''' <returns>実行結果（1つだけ返す。SQL文の集計関数等に有効）</returns>
    ''' <remarks></remarks>
    Public Function OleDBExcuteScalar(sqlCommand As String) As Object

        Dim cn As OleDb.OleDbConnection = New OleDb.OleDbConnection(_sqlConnect)
        Dim Com As OleDb.OleDbCommand = New OleDb.OleDbCommand(sqlCommand, cn)
        Dim RespObject As Object

        cn.Open()

        Try

            RespObject = Com.ExecuteScalar
            Return RespObject

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        Finally
            cn.Close()
        End Try

    End Function

End Class
