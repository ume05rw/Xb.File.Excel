Option Strict On

Imports System.Drawing

Partial Public Class Excel

    ''' <summary>
    ''' ユニットテストツール用のログ書き出しクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TestLogger
        Implements IDisposable

        Private _xls As Xb.File.Excel
        Private _index As Integer

        Private _fileName As String
        Private _directory As String
        Private _fullPath As String

        Private _db As Xb.Db.MsSql

        ''' <summary>
        ''' Dispose済みフラグ値を返す。
        ''' </summary>
        Public ReadOnly Property IsDisposed() As Boolean
            Get
                Return Me._disposedValue
            End Get
        End Property

        ''' <summary>
        ''' DBコネクション
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Db As Xb.Db.MsSql
            Get
                Return Me._db
            End Get
        End Property


        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="directory"></param>
        ''' <param name="dbName"></param>
        ''' <param name="userName"></param>
        ''' <param name="password"></param>
        ''' <param name="serverName"></param>
        Public Sub New(ByVal directory As String, _
                       ByVal dbName As String, _
                       Optional ByVal userName As String = "sa", _
                       Optional ByVal password As String = "sa", _
                       Optional ByVal serverName As String = "localhost")

            Me._db = New Xb.Db.MsSql(dbName, userName, password, serverName)
            Me._fileName = "TestLogger_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") & ".xls"

            If (Not System.IO.Directory.Exists(directory)) Then
                directory = "C:\log\"
            End If

            '実行ファイル
            'this._directory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            Me._directory = System.IO.Path.GetDirectoryName(directory)

            '出力先フォルダが存在しないとき、フォルダ作成を試みる
            If (Not System.IO.Directory.Exists(Me._directory)) Then
                Try
                    System.IO.Directory.CreateDirectory(Me._directory)
                    If (Not System.IO.Directory.Exists(Me._directory)) Then
                        Throw New Exception()
                    End If
                Catch ex As Exception
                    Throw New Exception("指定されたログ出力先フォルダ「" & Me._directory & "」を作成出来ませんでした。：" & ex.Message)
                End Try
            End If

            'ログファイルのフルパスを生成する。
            Me._fullPath = System.IO.Path.Combine(Me._directory, Me._fileName)

            Me._index = 1

            Try
                Me._xls = New File.Excel(Me._fullPath)
            Catch ex As Exception
                Throw New Exception("テスト結果書き出し用Excelファイルのオープンに失敗しました：" & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' ログを書き出す。
        ''' </summary>
        ''' <param name="text"></param>
        Public Sub Log(text As String)

            Dim logString As String = String.Format("{0} {1}", DateTime.Now.ToString("HH:mm:ss.fff"), text)

            Try
                Me._xls.SetValue("A" & Me._index.ToString(), logString)
                Me._index += 2
            Catch ex As Exception
                Console.WriteLine("ログ出力に失敗しました：" & logString)
            End Try
        End Sub

        ''' <summary>
        ''' DBへのクエリ結果をCSV化してログに書き出す。
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="wheres"></param>
        Public Sub LogQuery(tableName As String, Optional wheres As String() = Nothing)

            Dim where As String = ""
            Dim sql As String = "SELECT * FROM " & tableName
            If (Not wheres Is Nothing) Then
                For Each s As String In wheres
                    If (String.IsNullOrEmpty(s)) Then
                        If (where.Length > 0) Then
                            where &= " AND "
                        End If
                        where &= s
                    End If
                Next
                sql &= " WHERE " & where
            End If

            Me.Log("Query: " & sql)


            Dim dt As DataTable = Me._db.Query(sql)
            If (dt Is Nothing OrElse dt.Rows.Count <= 0) Then
                Me.Log("-- NO RESULT --")
            Else
                'カラム名を書き出す。
                Dim dtHeader As DataTable = New DataTable(), _
                    row As DataRow, _
                    endCell As String

                For i As Integer = 0 To dt.Columns.Count - 1
                    dtHeader.Columns.Add(i.ToString())
                Next

                row = dtHeader.NewRow()
                For i As Integer = 0 To dt.Columns.Count - 1
                    row(i) = dt.Columns(i).ColumnName
                Next
                dtHeader.Rows.Add(row)
                Me._xls.SetRangeTable("A" & Me._index.ToString(), dtHeader)

                'カラム名部を色付けする。
                endCell = File.Excel.GetColmunString(dtHeader.Columns.Count + 1) + Me._index.ToString()
                Me._xls.SetBackColor("A" & Me._index.ToString(), endCell, System.Drawing.Color.LightGray)

                Me._index += 1

                '明細を書き出す。
                Me._xls.SetRangeTable("A" & Me._index.ToString(), dt)

                'ヘッダ＋明細を枠線で囲む。
                endCell = File.Excel.GetColmunString(dtHeader.Columns.Count + 1) + (Me._index + dt.Rows.Count - 1).ToString()
                Me._xls.SetBorderLineAuto("A" & (Me._index - 1).ToString(), endCell)

                Me._index += dt.Rows.Count + 1

            End If


        End Sub

        ''' <summary>
        ''' 画像オブジェクトを保存する。
        ''' </summary>
        ''' <param name="image"></param>
        Public Sub LogImage(image As Image)

            Me._xls.SetImage("D" & Me._index.ToString(), image)
            Me._index += 4

        End Sub

        '''' <summary>
        '''' コントロールを画像ファイルとして保存する。
        '''' </summary>
        '''' <param name="control"></param>
        'Public Sub LogControl(control As Control)
        '    Dim bmp As New Bitmap(control.Width, control.Height)
        '    control.DrawToBitmap(bmp, New Rectangle(0, 0, control.Width, control.Height))

        '    Me.LogImage(bmp)
        'End Sub

    #Region "IDisposable Support"
        Private _disposedValue As Boolean ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me._disposedValue Then
                If disposing Then
                    Me._xls.Save()
                    Me._xls.Dispose()
                End If
            End If
            Me._disposedValue = True
        End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
    #End Region

    End Class

End Class


'DataGridViewを生成して画像化する。
'Dim value As String = File.Csv.GetCsvText(dt, Type.[String].LinefeedType.CrLf)

''DataGridViewにクエリ結果を張り付ける。
'Dim dgv As New DataGridView()
'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

''バインドしてもすぐ描画されるわけではないらしい。手動で書く。
''dgv.DataSource = dt;
''dgv.Refresh();

'Dim columnCount As Integer = 0
'For Each col As DataColumn In dt.Columns
'    dgv.Columns.Add(col.ColumnName, col.ColumnName)
'    columnCount += 1
'Next

'For i As Integer = 0 To dt.Rows.Count - 1
'    dgv.Rows.Add(1)
'    For j As Integer = 0 To columnCount - 1
'        dgv.Rows(i).Cells(j).Value = dt.Rows(i)(j)
'    Next
'Next

'Dim tmp As Integer = 0
'For Each col As DataGridViewColumn In dgv.Columns
'    tmp += col.Width
'Next
'dgv.Width = tmp

'tmp = 0
'For Each row As DataGridViewRow In dgv.Rows
'    tmp += row.Height
'Next
'dgv.Height = tmp

'Me.Log("Query: " & sql)
'Me.Log("Value: " & vbCr & vbLf & value)
'Me.LogControl(dgv)

'dgv.Dispose()
