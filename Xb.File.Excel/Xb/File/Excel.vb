Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Drawing

'デバッグ用記述：納品時は下記を使用しない。
'Microsoft.Excel.Objectへの参照設定も削除しておくこと
'Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' エクセルオブジェクト管理クラス
''' </summary>
''' <remarks>
''' 「遅延バインディングが出来ない」とエラーが出る。
''' Strict化は面倒なので断念。
''' 
''' Excelファイルにアクセスするには
''' http://www.atmarkit.co.jp/fdotnet/dotnettips/717excelfile/excelfile.html
''' </remarks>
Partial Public Class Excel
    Implements IDisposable


    ''' <summary>
    ''' シート上の方向区分
    ''' </summary>
    Public Enum EndPointDirection

        ''' <summary>
        ''' 右方向
        ''' </summary>
        ''' <remarks></remarks>
        Right

        ''' <summary>
        ''' 左方向
        ''' </summary>
        ''' <remarks></remarks>
        Left

        ''' <summary>
        ''' 上方向
        ''' </summary>
        ''' <remarks></remarks>
        Up

        ''' <summary>
        ''' 下方向
        ''' </summary>
        ''' <remarks></remarks>
        Down
    End Enum


    Private ReadOnly _path As String ' 管理対象Excelファイルのパス
    Private ReadOnly _process As Diagnostics.Process  ' Excelオブジェクトの実体プロセス
    Private _app As Object          'Excel.Application  Excelオブジェクト
    Private _book As Object         'Excel.Workbook     カレントワークブック
    Private _sheet As Object        'Excel.Worksheet    カレントシート
    Private _ranges As File.Excel.Range()   'Excel.Rangeオブジェクト配列。Close時の解放処理用。

    ''' <summary>
    ''' Excelファイルのパス
    ''' </summary>
    Public ReadOnly Property FileFullPath() As String
        Get
            Return Me._path
        End Get
    End Property

    ''' <summary>
    ''' Excelファイルのファイル名
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileName As String
        Get
            Return System.IO.Path.GetFileName(Me._path)
        End Get
    End Property

    ''' <summary>
    ''' Excelファイルの配置先パス
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Directory As String
        Get
            Return System.IO.Path.GetDirectoryName(Me._path)
        End Get
    End Property

    ''' <summary>
    ''' Excelのバージョンを取得する。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetVersion() As Integer

        'Excelオブジェクトを起動する。
        Try
            Dim app As Object = CreateObject("Excel.Application")
            Dim result As Integer = DirectCast(app.Version, Integer)

            Dim nullProcessId As Integer = 0
            GetWindowThreadProcessId(CInt(app.Hwnd), nullProcessId)

            Dim procs As Diagnostics.Process() = Diagnostics.Process.GetProcesses()
            Dim appProc As Diagnostics.Process = Nothing
            For Each proc As Diagnostics.Process In procs
                If (nullProcessId = proc.Id) Then
                    appProc = proc
                    Exit For
                End If
            Next

            'Excelオブジェクトを破棄する。
            app.DisplayAlerts = True
            app.Quit()

            'COMオブジェクトを解放する。
            Runtime.InteropServices.Marshal.ReleaseComObject(app)
            app = Nothing

            GC.Collect()

            If (appProc IsNot Nothing) Then
                appProc.Kill()
            End If
                
            Return result

        Catch ex As Exception
            Return -1
        End Try

    End Function



    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="path">管理対象のExcelファイルパス</param>
    Public Sub New(ByVal path As String)

        'Excelオブジェクトを起動する。
        Try
            Me._app = CreateObject("Excel.Application")
        Catch ex As Exception
            Xb.Util.Out("Util.Excel.New: Microsoft Excel がインストールされていません。")
            Throw New Exception("Microsoft Excel がインストールされていません。")
        End Try

        Me._app.Visible = False
        Me._app.DisplayAlerts = False

        'Excelアプリケーションのハンドルから、アプリケーションのプロセスを取得する。
        'close時にプロセスを破棄するため。
        Dim nullProcessId As Integer = 0
        GetWindowThreadProcessId(CInt(Me._app.Hwnd), nullProcessId)
        'Dim targetThreadId As Integer = GetWindowThreadProcessId(Me._app.Hwnd, nullProcessId)
        'ハンドルからプロセスIDを取得する
        '稼働中の全プロセスから、前述処理で生成したExcelアプリケーションに該当するプロセスを取得する。
        Dim procs As Diagnostics.Process() = Diagnostics.Process.GetProcesses()
        For Each proc As Diagnostics.Process In procs
            If (nullProcessId = proc.Id) Then
                Me._process = proc
                Exit For
            End If
        Next

        Me._path = path

        Me._ranges = New Excel.Range() {}

        '渡し値のExcelワークブックを、Excelオブジェクトで開く。
        Try
            If (Not IO.File.Exists(path)) Then
                'ファイルがないとき、新規作成する。
                Me._app.Workbooks.Add()
                Me._book = Me._app.ActiveWorkbook

                Dim sheets As Object = Me._book.Worksheets
                Dim idx As Integer
                sheets.Add()

                If (path.ToLower().IndexOf(".xls") = -1) Then path &= ".xls"

                If (Me._app.Version < 12) Then
                    'Excelバージョンが2003以前のとき
                    idx = path.ToLower().IndexOf(".xlsx")
                    If (idx <> -1) Then path = path.Substring(0, idx) & ".xls"
                Else
                    'Excelバージョンが2007移行のとき
                    path = path.Substring(0, path.ToLower().IndexOf(".xls")) & ".xlsx"
                End If

                '渡し値パスで新規Excelファイルを保存する。
                Me._book.SaveAs(path)
            Else
                'ファイルが存在するとき
                Me._book = Me._app.Workbooks.Open(path)
                ' UpdateLinks (0 / 1 / 2 / 3)
                ' ReadOnly (True / False )
                ' Format 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;) 5:なし / 6:引数 Delimiterで指定された文字
                ' Password
                ' WriteResPassword
                ' IgnoreReadOnlyRecommended
                ' Origin
                ' Delimiter
                ' Editable
                ' Notify
                ' Converter
                ' AddToMru
                ' Local
                ' CorruptLoad
            End If

        Catch ex As Exception
            Me.Close()
            Xb.Util.Out("Util.Excel.New: " & ex.Message & "：Excelファイル = " & path)
            Throw New Exception(ex.Message & "：Excelファイル = " & path)
        End Try

        'Excelファイル内の先頭シートを取得する。
        Try
            Me.SetCurrentSheet(1)
        Catch ex As Exception
            Me.Close()
            Xb.Util.Out("Util.Excel.New: " & ex.Message)
            Throw
        End Try
    End Sub


    ' ReSharper disable UnusedMethodReturnValue.Local
    ''' <summary>
    ''' ハンドルからプロセスIDを取得する。
    ''' </summary>
    ''' <param name="hWnd">ハンドル</param>
    ''' <param name="lpdwProcessId">プロセスID</param>
    ''' <returns></returns>
    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    End Function
    ' ReSharper restore UnusedMethodReturnValue.Local

    ''' <summary>
    ''' Excelブックオブジェクトの存在チェック
    ''' </summary>
    Private Sub CheckState()
        'Excelブックの存在チェック
        If ((Me._app Is Nothing) OrElse (Me._book Is Nothing)) Then
            Xb.Util.Out("Util.Excel.CheckState: Excelファイルがオープンされていません。")
            Throw New Exception("Excelファイルがオープンされていません。")
        End If
    End Sub


    ''' <summary>
    ''' セル文字列のフォーマットチェック
    ''' </summary>
    ''' <param name="cell"></param>
    Private Sub CheckLocation(ByVal cell As String)
        If (Not ValidateLocationFormat(cell)) Then
            Xb.Util.Out("Util.Excel.CheckLocation: セル文字列のフォーマットが異常です：" & cell)
            Throw New Exception("セル文字列のフォーマットが異常です：" & cell)
        End If
    End Sub


    ''' <summary>
    ''' シート名文字列から、シート番号を取得する。
    ''' </summary>
    ''' <param name="sheetName">シート名</param>
    ''' <returns>シート番号</returns>
    Public Function GetSheetIndex(ByVal sheetName As String) As Integer
        'Excelブックの存在チェック
        Me.CheckState()

        Dim idx As Integer = 1
        ' <- ブック内のシート番号なので、"1"始まり。
        For Each sh As Object In Me._book.Sheets
            If (sheetName = sh.Name.ToString()) Then
                Runtime.InteropServices.Marshal.ReleaseComObject(sh) 'COMオブジェクトを解放する。
                Return idx
            End If
            idx += 1
            Runtime.InteropServices.Marshal.ReleaseComObject(sh) 'COMオブジェクトを解放する。
        Next
        Xb.Util.Out("Util.Excel.GetSheetIndex: 渡し値シート名に該当するシートがありません：" & sheetName)
        Throw New Exception("渡し値シート名に該当するシートがありません：" & sheetName)
    End Function


    ''' <summary>
    ''' 本オブジェクト内で保持する、カレントシートをセットする。
    ''' </summary>
    ''' <param name="sheetName">シート名</param>
    Public Sub SetCurrentSheet(ByVal sheetName As String)
        'Excelブックの存在チェック
        Me.CheckState()

        Me._sheet = Me._book.Sheets(Me.GetSheetIndex(sheetName))
    End Sub


    ''' <summary>
    ''' 本オブジェクト内で保持する、カレントシートをセットする。
    ''' </summary>
    ''' <param name="index"></param>
    ''' <remarks></remarks>
    Public Sub SetCurrentSheet(ByVal index As Integer)
        'Excelブックの存在チェック
        Me.CheckState()

        If ((index < 1) OrElse (Me._book.Sheets.Count < index)) Then
            Xb.Util.Out("渡し値シート番号に該当するシートがありません：" & index)
            Throw New Exception("渡し値シート番号に該当するシートがありません：" & index)
        End If

        Me._sheet = Me._book.Sheets(index)
    End Sub


    'Excel.Sheetオブジェクトを渡す／取得する処理は不要になるように。
    ' ''' <summary>
    ' ''' 渡し値シート名、シート番号に合致するシートを返す。
    ' ''' </summary>
    ' ''' <param name="sheetName">シート名</param>
    ' ''' <returns>Excelシートオブジェクト</returns>
    'Public Function SetSheet(ByVal sheetName As String) As Object
    '    Me.SetCurrentSheet(sheetName)
    '    Return Me._sheet
    'End Function
    ' ''' <summary>
    ' ''' 渡し値シート名、シート番号に合致するシートを返す。
    ' ''' </summary>
    ' ''' <param name="index"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function GetSheet(ByVal index As Integer) As Object
    '    Me.SetCurrentSheet(index)
    '    Return Me._sheet
    'End Function


    ''' <summary>
    ''' ブックに含まれる全シート名を配列で取得する。
    ''' </summary>
    ''' <returns>シート名配列</returns>
    Public Function GetSheetNameArray() As String()
        'Excelブックの存在チェック
        Me.CheckState()

        Dim result As String() = New String(CInt(Me._book.Sheets.Count) - 1) {}

        Dim idx As Integer = 0
        ' <- シート番号でなく、配列のインデックスNo.なので、"0"始まり。
        For Each sh As Object In Me._book.Sheets
            result(idx) = sh.Name.ToString()
            Runtime.InteropServices.Marshal.ReleaseComObject(sh) 'COMオブジェクトを解放する。
            idx += 1
        Next

        Return result
    End Function


    ''' <summary>
    ''' 渡し値セルの値を文字型で取得する。
    ''' </summary>
    ''' <param name="cell">セル文字列</param>
    ''' <returns>セル内の値(文字型)</returns>
    Public Function GetValue(ByVal cell As String) As String
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(cell)

        '渡し値セル位置が異常のとき、空文字列を返す。
        If ((cell Is Nothing) OrElse (cell = "")) Then
            Return ""
        End If

        Try
            Dim rng As Object = Me._sheet.Range(cell)
            Dim result As String = rng.Text.ToString()
            Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
            Return result
        Catch ex As Exception
            Xb.Util.Out("Util.Excel.GetValue: セルの指定範囲を超えています： cell = " & cell)
            Throw New Exception("セルの指定範囲を超えています： cell = " & cell)
        End Try
    End Function


    ''' <summary>
    ''' 値を指定セルにセットする。
    ''' </summary>
    ''' <param name="cell"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal cell As String, ByVal value As String)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(cell)

        Try
            Dim rng As Object = Me._sheet.Range(cell)
            Dim tryValue As Decimal
            If (Not Convert.IsDBNull(value)) Then
                If (Decimal.TryParse(value.ToString(), tryValue)) Then
                    rng.Value2 = tryValue
                Else
                    rng.Value2 = value
                End If
            Else
                rng.Value2 = String.Empty
            End If

            Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
        Catch ex As Exception
            Xb.Util.Out("Util.Excel.SetValue: セルの指定範囲を超えています： cell = " & cell)
            Throw New Exception("セルの指定範囲を超えています： cell = " & cell)
        End Try
    End Sub


    ''' <summary>
    ''' Rangeオブジェクトを取得する。
    ''' Close/Dispose時にオブジェクトを一括解放するため、Object配列に保持しておく。
    ''' </summary>
    ''' <param name="fromCell"></param>
    ''' <param name="toCell"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRange(ByVal fromCell As String, ByVal toCell As String) As Range
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        If (Not IsArray(Me._ranges)) Then Me._ranges = New Range() {}
        ReDim Preserve Me._ranges(Me._ranges.Length)

        Me._ranges(Me._ranges.Length - 1) = New Range(Me._sheet.Range(fromCell, toCell))

        Return Me._ranges(Me._ranges.Length - 1)
    End Function


    ''' <summary>
    ''' 2セル間の範囲の値をDataTableにセットして返す。
    ''' </summary>
    ''' <param name="fromCell">範囲開始セル位置</param>
    ''' <param name="toCell">範囲終了セル位置</param>
    ''' <returns>セル範囲のデータ入りDataTable</returns>
    Public Function GetRangeTable(ByVal fromCell As String, ByVal toCell As String) As DataTable
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        '応答値DataTableを初期化する。
        Dim result As New DataTable()

        '渡し値範囲を取得する。ｌ
        Dim rng As Object = Me._sheet.Range(fromCell, toCell)

        '取得値が配列か否かで分岐する。
        '複数セルが指定されている場合は、必ず2次元配列にキャスト可能。
        'fromCell, toCellが同じセルのとき、配列にならない。
        ' ReSharper disable VBPossibleMistakenCallToGetType.2
        If (rng.Value2.[GetType]().IsArray) Then
            ' ReSharper restore VBPossibleMistakenCallToGetType.2
            '範囲の値が配列のとき
            '(object[,])rng.Value2 では、テキストでなく内部的に保持している値を取得してしまう。
            '日付が表示されているとき、double型値として取得されたりなど。
            Dim values As Object(,) = DirectCast(rng.Value(System.Type.Missing), Object(,))

            Dim iRow As Integer = values.GetLength(0)
            Dim iCol As Integer = values.GetLength(1)

            '取得列数分、カラムを追加する。
            For j As Integer = 0 To iCol - 1
                result.Columns.Add(System.String.Format("F{0}", (j + 1)))
            Next

            'DataTableへ値をセットする。
            For i As Integer = 0 To iRow - 1
                Dim dr As DataRow = result.NewRow()
                For j As Integer = 0 To iCol - 1
                    dr(j) = If(values((i + 1), (j + 1)) Is Nothing, "", values((i + 1), (j + 1)).ToString())
                Next
                result.Rows.Add(dr)
            Next
        Else
            '範囲の値が単一のとき
            result.Columns.Add("F1")
            Dim dr As DataRow = result.NewRow()
            dr(0) = rng.Value(System.Type.Missing).ToString()
            result.Rows.Add(dr)
        End If

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
        Return result
    End Function


    ''' <summary>
    ''' 指定セルを起点に、DataTable内の値を貼り付ける
    ''' </summary>
    ''' <param name="startCell"></param>
    ''' <param name="dt"></param>
    Public Sub SetRangeTable(ByVal startCell As String, ByRef dt As DataTable)
        'セル文字列フォーマットチェック
        Me.CheckLocation(startCell)

        If (dt Is Nothing) Then
            Xb.Util.Out("Util.Excel.SetRangeTable: DataTableがnullです。")
            Throw New Exception("DataTableがnullです。")
        End If

        '1000行を超えるデータを張り付けるとき、1000行ごとに区切って貼り付け処理を行う。
        If (dt.Rows.Count > 1000) Then
            Dim dtClone As DataTable = dt.Clone()
            Dim startColumnMark As String = Excel.GetColumnStringByLocation(startCell)
            Dim startRowIndex As Integer = Excel.GetRowIndexByLocation(startCell)

            dtClone.Rows.Clear()

            For iRow As Integer = 0 To dt.Rows.Count - 1
                Dim cloneRow As DataRow = dtClone.NewRow()
                For iCol As Integer = 0 To dt.Columns.Count - 1
                    cloneRow(iCol) = dt.Rows(iRow)(iCol)
                Next
                dtClone.Rows.Add(cloneRow)

                If (dtClone.Rows.Count >= 1000 _
                    OrElse iRow = dt.Rows.Count - 1) Then

                    Me.SetRangeTable(startColumnMark & startRowIndex.ToString(), dtClone)
                    startRowIndex += dtClone.Rows.Count
                    dtClone.Rows.Clear()
                End If
            Next

            Return
        End If


        '終点セルの位置文字列を取得する。
        Dim iStartColumnIndex As Integer = Excel.GetColumnIndex(Excel.GetColumnStringByLocation(startCell))
        Dim iStartRowIndex As Integer = Excel.GetRowIndexByLocation(startCell)
        Dim sEndCell As String = System.String.Format("{0}{1}", Excel.GetColmunString(iStartColumnIndex + dt.Columns.Count - 1), (iStartRowIndex + dt.Rows.Count - 1).ToString())

        '値の貼り付け領域を取得する。
        Dim rng As Object = Me._sheet.Range(startCell, sEndCell)

        '貼り付け用2次元配列を生成する。
        Dim data As Object(,) = New Object(dt.Rows.Count - 1, dt.Columns.Count - 1) {}
        Dim tryValue As Decimal
        For iRow As Integer = 0 To dt.Rows.Count - 1
            For iCol As Integer = 0 To dt.Columns.Count - 1
                If (Not Convert.IsDBNull(dt.Rows(iRow)(iCol))) Then

                    'Console.WriteLine("ColumnName: " & dt.Columns.Item(iCol).ColumnName & "  /  DataType: " & dt.Columns.Item(iCol).DataType.ToString())

                    If (Decimal.TryParse(dt.Rows(iRow)(iCol).ToString(), tryValue)) Then
                        '数値にキャスト可能な値のとき

                        If (dt.Columns.Item(iCol).DataType.ToString() = "System.String") Then
                            'DataTable上の型が文字列のとき、文字列型で出力
                            data(iRow, iCol) = "'" & dt.Rows(iRow)(iCol).ToString()
                        Else
                            'DataTable上の型が文字列でないとき、数値型で出力
                            data(iRow, iCol) = tryValue
                        End If
                    Else
                        data(iRow, iCol) = "'" & dt.Rows(iRow)(iCol).ToString()
                    End If
                Else
                    data(iRow, iCol) = String.Empty
                End If
            Next
        Next

        '範囲へ値を貼り付ける。
        rng.Value2 = data

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub


    ''' <summary>
    ''' 指定セルから右方向へ、渡し値文字列配列を順次セットする。
    ''' </summary>
    ''' <param name="startCell">セル開始位置</param>
    ''' <param name="values">書き込み用値配列</param>
    Public Sub SetRowValues(ByVal startCell As String, ByVal values As String())
        'セル文字列フォーマットチェック
        Me.CheckLocation(startCell)

        Dim row As Integer = Excel.GetRowIndexByLocation(startCell)
        Dim toCellColumn As String = Excel.GetColmunString(Excel.GetColumnIndex(Excel.GetColumnStringByLocation(startCell)) + values.Length - 1)

        '値の貼り付け領域を取得する。
        Dim rng As Object = Me._sheet.Range(startCell, System.String.Format("{0}{1}", toCellColumn, row.ToString()))

        '貼り付け用2次元配列を生成する。
        Dim data As String(,) = New String(0, values.Length - 1) {}
        For i As Integer = 0 To values.Length - 1
            If (values(i) IsNot Nothing) Then
                data(0, i) = values(i)
            Else
                data(0, i) = ""
            End If
        Next

        rng.Value2 = data

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub


    ''' <summary>
    ''' 指定セルに画像を貼り付ける。
    ''' </summary>
    ''' <param name="objImage"></param>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Public Sub SetImage(ByVal cell As String, _
                        ByRef objImage As Drawing.Image, _
                        Optional ByVal marginLeft As Integer = 0, _
                        Optional ByVal marginTop As Integer = 0)

        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(cell)

        '渡し値の画像オブジェクトをテンポラリ領域に保存する。
        Dim tmpImageFileName As String = IO.Path.GetTempFileName()
        objImage.Save(tmpImageFileName, Drawing.Imaging.ImageFormat.Png)

        Dim rng As Object = Me._sheet.Range(cell)
        Dim sngLeft As Single
        Dim sngTop As Single
        Dim sngWidth As Single
        Dim sngHeight As Single

        Single.TryParse((rng.Left + marginLeft).ToString(), sngLeft)
        Single.TryParse((rng.Top + marginTop).ToString(), sngTop)
        Single.TryParse(objImage.Width.ToString(), sngWidth)
        Single.TryParse(objImage.Height.ToString(), sngHeight)

        'Microsoft.Office.Core.MsoTriState.msoFalse = 0
        'Microsoft.Office.Core.MsoTriState.msoTrue = -1
        Me._sheet.Shapes.AddPicture(tmpImageFileName, _
                                    0, _
                                    -1, _
                                    sngLeft, _
                                    sngTop, _
                                    sngWidth * 0.9, _
                                    sngHeight * 0.9)

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)

        IO.File.Delete(tmpImageFileName)

    End Sub


    ''' <summary>
    ''' 渡し値Rangeオブジェクトを、指定セルを基準位置にしてコピーする。
    ''' </summary>
    ''' <param name="cell">コピー基準セル位置</param>
    ''' <param name="objRange">コピー元Range</param>
    ''' <remarks></remarks>
    Public Sub CopyRange(ByVal cell As String, ByRef objRange As Range)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(cell)

        'Rangeオブジェクトの存在チェック
        If (objRange Is Nothing) Then
            Xb.Util.Out("Util.Excel.CopyRange_1: Rangeオブジェクトが検出できません。")
            Throw New Exception("Rangeオブジェクトが検出できません。")
        End If

        'クリップボードに渡し値Rangeオブジェクトをコピーする。
        objRange.Value.Copy()
        Dim rng As Object = Me._sheet.Range(cell)
        Try
            rng.PasteSpecial()
        Catch ex As Exception
            Xb.Util.Out("Util.Excel.CopyRange_2: Rangeコピーに失敗しました。：" & ex.Message)
            Throw New Exception("Rangeコピーに失敗しました。：" & ex.Message)
        End Try

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub


    ''' <summary>
    ''' 渡し値Rangeオブジェクトを、指定セルを基準位置にしてコピーする。
    ''' </summary>
    ''' <param name="fromCell"></param>
    ''' <param name="toCell"></param>
    ''' <param name="objRange"></param>
    ''' <remarks></remarks>
    Public Sub CopyRange(ByVal fromCell As String, ByVal toCell As String, ByRef objRange As Range)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        'Rangeオブジェクトの存在チェック
        If (objRange Is Nothing) Then
            Xb.Util.Out("Util.Excel.CopyRange_3: Rangeオブジェクトが検出できません。")
            Throw New Exception("Rangeオブジェクトが検出できません。")
        End If

        'クリップボードに渡し値Rangeオブジェクトをコピーする。
        objRange.Value.Copy()
        Dim rng As Object = Me._sheet.Range(fromCell, toCell)
        Try
            rng.PasteSpecial()
        Catch ex As Exception
            Xb.Util.Out("Util.Excel.CopyRange_3: Rangeコピーに失敗しました。：" & ex.Message)
            Throw New Exception("Rangeコピーに失敗しました。：" & ex.Message)
        End Try

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub

    ''' <summary>
    ''' 指定範囲を選択状態にする。
    ''' </summary>
    ''' <param name="fromCell"></param>
    ''' <param name="toCell"></param>
    ''' <remarks></remarks>
    Public Sub [Select](ByVal fromCell As String, _
                        ByVal toCell As String)

        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        Dim rng As Object = Me._sheet.Range(fromCell, toCell)
        rng.Select()

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub

    ''' <summary>
    ''' 指定セルを選択状態にする。
    ''' </summary>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Public Sub [Select](ByVal cell As String)
        Me.Select(cell, cell)
    End Sub


    ''' <summary>
    ''' 指定範囲の背景色をセットする。
    ''' </summary>
    ''' <param name="fromCell">範囲開始セル位置</param>
    ''' <param name="toCell">範囲終了セル位置</param>
    ''' <param name="color"></param>
    Public Sub SetBackColor(ByVal fromCell As String, _
                            ByVal toCell As String, _
                            ByVal color As System.Drawing.Color)

        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        Dim rng As Object = Me._sheet.Range(fromCell, toCell)
        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color)

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub

    ''' <summary>
    ''' 指定セルの背景色をセットする。
    ''' </summary>
    ''' <param name="cell"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetBackColor(ByVal cell As String, ByVal color As System.Drawing.Color)
        Me.SetBackColor(cell, cell, color)
    End Sub


    ''' <summary>
    ''' 指定範囲の文字色をセットする。
    ''' </summary>
    ''' <param name="fromCell"></param>
    ''' <param name="toCell"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetForeColor(ByVal fromCell As String, _
                            ByVal toCell As String, _
                            ByVal color As System.Drawing.Color)

        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        Dim rng As Object = Me._sheet.Range(fromCell, toCell)
        rng.Font.Color = System.Drawing.ColorTranslator.ToOle(color)

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub

    ''' <summary>
    ''' 指定セルの文字色をセットする。
    ''' </summary>
    ''' <param name="cell"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetForeColor(ByVal cell As String, _
                            ByVal color As System.Drawing.Color)
        Me.SetForeColor(cell, cell, color)
    End Sub

    ''' <summary>
    ''' 指定セルに数式をセットする。
    ''' </summary>
    ''' <param name="cell"></param>
    ''' <param name="formula"></param>
    ''' <remarks></remarks>
    Public Sub SetFormula(ByVal cell As String, _
                            ByVal formula As String)

        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(cell)

        Dim rng As Object = Me._sheet.Range(cell, cell)
        rng.Formula = formula

    End Sub


    Public Enum BorderType
        Top = 8
        Bottom = 9
        Left = 7
        Right = 10
        InsideHorizontal = 12
        InsideVertical = 11

        '以下はExcel定数ではない。
        TopAndBottom = 100
        LeftAndRight = 101
        OutsideAll = 102
        InsideAll = 103
    End Enum

    Public Enum LineStyle
        None = -4142
        Normal = 1
        Dash = -4115
        DashDot = 4
        DashDotDot = 5
        Dot = -4118
        [Double] = -4119
        SlantDashDot = 13
    End Enum

    Public Enum LineWeight
        Tiny = 1
        Normal = 2
        Medium = -4138
        Thick = 4
    End Enum

    ''' <summary>
    ''' 指定範囲の枠線をセットする。
    ''' </summary>
    ''' <param name="fromCell">範囲開始セル位置</param>
    ''' <param name="toCell">範囲終了セル位置</param>
    Public Sub SetBorderLineAuto(ByVal fromCell As String, ByVal toCell As String)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        Dim rng As Object = Me._sheet.Range(fromCell, toCell)

        rng.Borders.Color = ColorTranslator.ToOle(Color.Black)
        rng.Borders(CInt(BorderType.Top)).LineStyle = CInt(LineStyle.Normal)
        rng.Borders(CInt(BorderType.Top)).Weight = CInt(LineWeight.Normal)

        rng.Borders(CInt(BorderType.Bottom)).LineStyle = CInt(LineStyle.Normal)
        rng.Borders(CInt(BorderType.Bottom)).Weight = CInt(LineWeight.Normal)

        rng.Borders(CInt(BorderType.Left)).LineStyle = CInt(LineStyle.Normal)
        rng.Borders(CInt(BorderType.Left)).Weight = CInt(LineWeight.Normal)

        rng.Borders(CInt(BorderType.Right)).LineStyle = CInt(LineStyle.Normal)
        rng.Borders(CInt(BorderType.Right)).Weight = CInt(LineWeight.Normal)

        rng.Borders(CInt(BorderType.InsideHorizontal)).LineStyle = CInt(LineStyle.Normal)
        rng.Borders(CInt(BorderType.InsideHorizontal)).Weight = CInt(LineWeight.Tiny)

        rng.Borders(CInt(BorderType.InsideVertical)).LineStyle = CInt(LineStyle.Normal)
        rng.Borders(CInt(BorderType.InsideVertical)).Weight = CInt(LineWeight.Tiny)

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub

    ''' <summary>
    ''' 指定範囲の枠線を詳細にセットする。
    ''' </summary>
    ''' <param name="fromCell"></param>
    ''' <param name="toCell"></param>
    ''' <param name="borderType"></param>
    ''' <param name="lineStyle"></param>
    ''' <param name="lineWeight"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetBorderLine(ByVal fromCell As String, _
                                ByVal toCell As String, _
                                ByVal borderType As BorderType, _
                                Optional ByVal lineStyle As LineStyle = LineStyle.Normal, _
                                Optional ByVal lineWeight As LineWeight = LineWeight.Normal, _
                                Optional ByVal color As Color = Nothing)

        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        If (color = Nothing) Then
            color = color.Black
        End If

        Dim rng As Object = Me._sheet.Range(fromCell, toCell)

        If (borderType = borderType.TopAndBottom) Then

            rng.Borders(CInt(borderType.Top)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Top)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Top)).Weight = CInt(lineWeight)

            rng.Borders(CInt(borderType.Bottom)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Bottom)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Bottom)).Weight = CInt(lineWeight)

        ElseIf (borderType = borderType.LeftAndRight) Then

            rng.Borders(CInt(borderType.Left)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Left)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Left)).Weight = CInt(lineWeight)

            rng.Borders(CInt(borderType.Right)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Right)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Right)).Weight = CInt(lineWeight)

        ElseIf (borderType = borderType.OutsideAll) Then

            rng.Borders(CInt(borderType.Top)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Top)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Top)).Weight = CInt(lineWeight)

            rng.Borders(CInt(borderType.Bottom)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Bottom)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Bottom)).Weight = CInt(lineWeight)

            rng.Borders(CInt(borderType.Left)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Left)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Left)).Weight = CInt(lineWeight)

            rng.Borders(CInt(borderType.Right)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.Right)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.Right)).Weight = CInt(lineWeight)

        ElseIf (borderType = borderType.InsideAll) Then

            rng.Borders(CInt(borderType.InsideHorizontal)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.InsideHorizontal)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.InsideHorizontal)).Weight = CInt(lineWeight)

            rng.Borders(CInt(borderType.InsideVertical)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType.InsideVertical)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType.InsideVertical)).Weight = CInt(lineWeight)

        Else
            rng.Borders(CInt(borderType)).LineStyle = CInt(lineStyle)
            rng.Borders(CInt(borderType)).Color = ColorTranslator.ToOle(color)
            rng.Borders(CInt(borderType)).Weight = CInt(lineWeight)
        End If




        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。

    End Sub

    ''' <summary>
    ''' 指定セルの枠線を詳細にセットする。
    ''' </summary>
    ''' <param name="cell"></param>
    ''' <param name="borderType"></param>
    ''' <param name="lineStyle"></param>
    ''' <param name="lineWeight"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetBorderLine(ByVal cell As String, _
                                ByVal borderType As BorderType, _
                                Optional ByVal lineStyle As LineStyle = LineStyle.Normal, _
                                Optional ByVal lineWeight As LineWeight = LineWeight.Normal, _
                                Optional ByVal color As Color = Nothing)
        Me.SetBorderLine(cell, cell, borderType, lineStyle, lineWeight, color)
    End Sub


    ''' <summary>
    ''' 指定列の幅をセットする。
    ''' </summary>
    ''' <param name="fromColumn">範囲開始列位置</param>
    ''' <param name="toColumn">範囲終了列位置</param>
    ''' <param name="width">幅ピクセル</param>
    Public Sub SetColumnWidth(ByVal fromColumn As String, ByVal toColumn As String, ByVal width As Integer)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromColumn & "1")
        Me.CheckLocation(toColumn & "1")

        Dim rng As Object = Me._sheet.Range(fromColumn & "1", toColumn & "1")
        rng.ColumnWidth = width
        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub


    ''' <summary>
    ''' 指定列の幅をセットする。
    ''' </summary>
    ''' <param name="column"></param>
    ''' <param name="width"></param>
    ''' <remarks></remarks>
    Public Sub SetColumnWidth(ByVal column As String, ByVal width As Integer)
        Me.SetColumnWidth(column, column, width)
    End Sub


    ''' <summary>
    ''' 指定行の高さをセットする。
    ''' </summary>
    ''' <param name="fromRow">範囲開始行位置</param>
    ''' <param name="toRow">範囲終了行位置</param>
    ''' <param name="height">高さピクセル</param>
    ''' <remarks></remarks>
    Public Sub SetRowHeight(ByVal fromRow As Integer, ByVal toRow As Integer, ByVal height As Integer)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation("A" & fromRow.ToString())
        Me.CheckLocation("A" & toRow.ToString())

        Dim rng As Object = Me._sheet.Rows(fromRow.ToString() & ":" & toRow.ToString())
        'Dim rng As Object = Me._sheet.Range("A" & fromRow.ToString(), "A" & toRow.ToString())

        rng.RowHeight = height
        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub


    ''' <summary>
    ''' 指定行の高さをセットする。
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub SetRowHeight(ByVal row As Integer, ByVal height As Integer)
        Me.SetRowHeight(row, row, height)
    End Sub


    ''' <summary>
    ''' 指定範囲の列幅を自動調整する。
    ''' </summary>
    ''' <param name="fromColumn">範囲開始列位置</param>
    ''' <param name="toColumn">範囲終了列位置</param>
    Public Sub AutoFitColumnWidth(ByVal fromColumn As String, ByVal toColumn As String)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromColumn & "1")
        Me.CheckLocation(toColumn & "1")


        Dim rndPoint As String = Me.GetEndPoint(fromColumn & "65536", File.Excel.EndPointDirection.Up)
        Dim rowIndex As Integer = File.Excel.GetRowIndexByLocation(rndPoint)
        Dim rng As Object = Me._sheet.Range(fromColumn & "1", toColumn & rowIndex.ToString())

        rng.Columns.AutoFit()
        'If (Me._app.Version < 12) Then

        'Else
        '    rng.EntireColumn.AutoFit()
        'End If

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub


    ''' <summary>
    ''' 指定範囲の列幅を自動調整する。
    ''' </summary>
    ''' <param name="column"></param>
    ''' <remarks></remarks>
    Public Sub AutoFitColumnWidth(ByVal column As String)
        Me.AutoFitColumnWidth(column, column)
    End Sub


    ''' <summary>
    ''' 指定範囲の行幅を自動調整する。
    ''' </summary>
    ''' <param name="fromRow">範囲開始行位置</param>
    ''' <param name="toRow">範囲終了行位置</param>
    ''' <remarks></remarks>
    Public Sub AutoFitRowHeight(ByVal fromRow As Integer, ByVal toRow As Integer)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation("A" & fromRow.ToString())
        Me.CheckLocation("A" & toRow.ToString())
        Dim rng As Object = Me._sheet.Range("A" & fromRow.ToString(), "A" & toRow.ToString())
        rng.AutoFit()
        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。

    End Sub


    ''' <summary>
    ''' 指定範囲の行幅を自動調整する。
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Public Sub AutoFitRowHeight(ByVal row As Integer)
        Me.AutoFitRowHeight(row, row)
    End Sub


    ''' <summary>
    ''' 値入力状態、もしくは空白状態が連続したセルの、末尾セル位置を取得する。
    ''' ※Excelシート上で、Ctrl + 方向キーを押したときの挙動と同じ。
    ''' </summary>
    ''' <param name="startCell">基準セル位置</param>
    ''' <param name="direction">方向</param>
    ''' <returns>セル位置文字列</returns>
    Public Function GetEndPoint(ByVal startCell As String, ByVal direction As EndPointDirection) As String
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(startCell)

        Dim objDirection As Integer
        Select Case direction
            Case EndPointDirection.Down
                objDirection = -4121    'Microsoft.Office.Interop.Excel.XlDirection.xlDown
                Exit Select
            Case EndPointDirection.Left
                objDirection = -4159    'Microsoft.Office.Interop.Excel.XlDirection.xlToLeft
                Exit Select
            Case EndPointDirection.Right
                objDirection = -4161    'Microsoft.Office.Interop.Excel.XlDirection.xlToRight
                Exit Select
            Case EndPointDirection.Up
                objDirection = -4162    'Microsoft.Office.Interop.Excel.XlDirection.xlUp
                Exit Select
        End Select


        Dim rng As Object = Me._sheet.Range(startCell).End(objDirection)

        Dim result As String = Excel.GetColmunString(CType(rng.Column, Integer)) & rng.Row.ToString()

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
        Return result
    End Function


    ''' <summary>
    ''' ブックを保存する。
    ''' </summary>
    Public Sub Save()
        If (Me._book IsNot Nothing) Then
            Me._book.Save()
        End If
    End Sub


    ''' <summary>
    ''' 縦方向の改ページを挿入する。
    ''' </summary>
    ''' <param name="row">改ページする行番号</param>
    ''' <remarks></remarks>
    Public Sub AddPageBreak(ByVal row As Integer)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation("A" & row.ToString())

        Dim rng As Object = Me._sheet.Range("A" & row.ToString())

        Me._sheet.HPageBreaks.Add(rng)

        Runtime.InteropServices.Marshal.ReleaseComObject(rng)    'COMオブジェクトを解放する。
    End Sub


    ''' <summary>
    ''' 印刷範囲を指定する。
    ''' </summary>
    ''' <param name="fromCell"></param>
    ''' <param name="toCell"></param>
    ''' <remarks></remarks>
    Public Sub SetPrintArea(ByVal fromCell As String, ByVal toCell As String)
        'Excelブックの存在チェック
        Me.CheckState()

        'セル文字列フォーマットチェック
        Me.CheckLocation(fromCell)
        Me.CheckLocation(toCell)

        Me._sheet.PageSetup.PrintArea = String.Format( _
            "${0}${1}:${2}${3}", _
            Excel.GetColumnStringByLocation(fromCell), _
            Excel.GetRowIndexByLocation(fromCell).ToString(), _
            Excel.GetColumnStringByLocation(toCell), _
            Excel.GetRowIndexByLocation(toCell).ToString() _
            )

    End Sub


    ''' <summary>
    ''' ブックの再計算を実行する。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Calculate()
        Me._app.Calculate()
    End Sub

    ''' <summary>
    ''' ブックの自動計算を開始／停止する。
    ''' </summary>
    ''' <param name="isDoAuto"></param>
    ''' <remarks></remarks>
    Public Sub SetAutoCalculation(ByVal isDoAuto As Boolean)

        'xlCalculationAutomatic = -4105
        'xlCalculationManual    = -4135
        Me._app.Calculation = If(isDoAuto, -4105, -4135)

    End Sub

    ''' <summary>
    ''' ブックの描画処理を開始／停止する。
    ''' </summary>
    ''' <param name="isUpdate"></param>
    ''' <remarks></remarks>
    Public Sub SetScreenUpdate(ByVal isUpdate As Boolean)

        Me._app.ScreenUpdating = isUpdate

    End Sub


    ''' <summary>
    ''' カレントシートを印刷する。
    ''' </summary>
    ''' <param name="printerName"></param>
    ''' <remarks></remarks>
    Public Sub Print(Optional ByVal printerName As String = "DEFAULT_PRINTER")
        'Excelブックの存在チェック
        Me.CheckState()

        If (printerName = "DEFAULT_PRINTER") Then
            Me._sheet.PrintOut(Copies:=1, Collate:=True)
        Else
            Me._sheet.PrintOut(Copies:=1, ActivePrinter:=printerName, Collate:=True)
        End If

    End Sub


    ''' <summary>
    ''' ブック、Excelオブジェクトを閉じる。
    ''' </summary>
    Public Sub Close()
        Me.Dispose()
    End Sub


    ''' <summary>
    ''' Excelプロセスを破棄する。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub KillExcelProcess(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs)
        Try
            Me._process.Kill()
        Catch ex As Exception
        End Try
    End Sub


    ''' <summary>
    ''' 指定座標(X/Yとも0始まり)のセル位置文字列を取得する。
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLocation(Optional ByVal x As Integer = 0, _
                                        Optional ByVal y As Integer = 0) As String

        Return File.Excel.GetColmunString(x + 1) & (y + 1).ToString()

    End Function


    ''' <summary>
    ''' 指定セルから、指定増分座標を移動した先のセル位置文字列を取得する。
    ''' </summary>
    ''' <param name="baseCell"></param>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLocation(ByVal baseCell As String, _
                                        Optional ByVal x As Integer = 0, _
                                        Optional ByVal y As Integer = 0) As String

        Dim baseX As Integer _
            = File.Excel.GetColumnIndex(File.Excel.GetColumnStringByLocation(baseCell)), _
            baseY As Integer _
            = File.Excel.GetRowIndexByLocation(baseCell), _
            targetX As Integer = baseX + x, _
            targetY As Integer = baseY + y

        Return File.Excel.GetLocation(targetX - 1, targetY - 1)

    End Function


    ''' <summary>
    ''' 列番号を文字列化する。(xls形式のみ対応。)
    ''' </summary>
    ''' <param name="col">列番号</param>
    ''' <returns>列位置文字列</returns>
    Public Shared Function GetColmunString(ByVal col As Integer) As String
        Dim result As String = ""

        Dim digit As Integer = 0
        Dim colRemnant As Integer = col
        Dim alp As Integer

        Do
            alp = CInt((colRemnant - 1) Mod 26)
            result = BaseStr.Substring(alp + 1, 1) & result
            colRemnant = CInt((colRemnant - alp) / 26)
        Loop While (colRemnant <> 0)

        Return result
    End Function


    Private Const BaseStr As String = " ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    ''' <summary>
    ''' 列文字列を番号化する。(xls形式のみ対応。)
    ''' </summary>
    ''' <param name="col">列位置文字列</param>
    ''' <returns>列番号</returns>
    Public Shared Function GetColumnIndex(ByVal col As String) As Integer

        col = col.ToUpper()

        If (col = "NULL") Then
            Return -1
        End If

        Dim result As Integer = 0
        Dim colChar As String

        For i As Integer = 0 To col.Length - 1
            colChar = col.Substring((col.Length - i - 1), 1)
            result += (BaseStr.IndexOf(col.Substring((col.Length - i - 1), 1))) * (26 ^ i)
        Next

        Return result

    End Function


    ''' <summary>
    ''' セル文字列から列位置文字列を取得する。
    ''' </summary>
    ''' <param name="cell">セル位置文字列</param>
    ''' <returns>列位置文字列</returns>
    Public Shared Function GetColumnStringByLocation(ByVal cell As String) As String
        If (cell Is Nothing) Then
            cell = ""
        End If
        Dim reg As New System.Text.RegularExpressions.Regex("[0-9]*")
        Return reg.Replace(cell, "").ToUpper()
    End Function


    ''' <summary>
    ''' セル文字列から行番号を取得する。
    ''' </summary>
    ''' <param name="cell">セル位置文字列</param>
    ''' <returns>行番号</returns>
    Public Shared Function GetRowIndexByLocation(ByVal cell As String) As Integer
        Dim reg As New System.Text.RegularExpressions.Regex("[0-9]*")
        Dim row As Integer = -1
        For Each mat As System.Text.RegularExpressions.Match In reg.Matches(cell)
            If (mat.Value.Length > 0) Then
                row = Integer.Parse(mat.Value)
                Exit For
            End If
        Next
        Return row
    End Function


    ''' <summary>
    ''' セル文字列のフォーマットを検証する。
    ''' </summary>
    ''' <param name="cell">セル位置文字列</param>
    ''' <returns>検証結果Boolean</returns>
    Public Shared Function ValidateLocationFormat(ByVal cell As String) As Boolean
        If (Excel.GetRowIndexByLocation(cell) = -1) Then
            Return False
        End If
        Dim colStr As String = Excel.GetColumnStringByLocation(cell)
        If colStr = "" Then
            Return False
        End If

        'Excel2003までの列制限
        'Dim colIdx As Integer = Excel.GetColumnIndex(colStr)
        'If ((0 >= colIdx) OrElse (colIdx > 256)) Then
        '    Return False
        'End If

        Return True
    End Function


    Private _disposedValue As Boolean ' 重複する呼び出しを検出するには

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me._disposedValue Then
            If disposing Then
                'Rangeオブジェクト配列の要素を解放する。
                If (IsArray(Me._ranges)) Then
                    For Each obj As Range In Me._ranges
                        Try
                            obj.Dispose()
                        Catch ex As Exception
                            Xb.Util.Out(ex)
                        End Try
                    Next
                End If

                'COMオブジェクトを解放する。
                If (Me._sheet IsNot Nothing) Then
                    Try
                        Runtime.InteropServices.Marshal.ReleaseComObject(Me._sheet)
                    Catch ex As Exception
                        Xb.Util.Out(ex)
                    End Try
                End If

                'ブックを閉じる
                If (Me._book IsNot Nothing) Then
                    Try
                        Me._book.Close(False, System.Type.Missing, False)
                    Catch ex As Exception
                        Xb.Util.Out(ex)
                    End Try
                End If

                'Excelオブジェクトを破棄する。
                If (Me._app IsNot Nothing) Then
                    Try
                        Me._app.DisplayAlerts = True
                        Me._app.Quit()
                    Catch ex As Exception
                        Xb.Util.Out(ex)
                    End Try

                    'COMオブジェクトを解放する。
                    Try
                        Runtime.InteropServices.Marshal.ReleaseComObject(Me._app)
                    Catch ex As Exception
                        Xb.Util.Out(ex)
                    End Try
                End If

                Me._sheet = Nothing
                Me._book = Nothing
                Me._app = Nothing

                GC.Collect()

                If (Me._process IsNot Nothing) Then
                    Dim timer As Timers.Timer = New Timers.Timer()
                    AddHandler timer.Elapsed, New Timers.ElapsedEventHandler(AddressOf Me.KillExcelProcess)
                    timer.Interval = 5000
                    timer.AutoReset = False
                    timer.Start()
                End If
            End If
        End If
        Me._disposedValue = True
    End Sub

#Region "IDisposable Support"
    ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
