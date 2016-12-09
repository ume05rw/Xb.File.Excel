Option Strict On

'デバッグ用記述：納品時は下記を使用しない。
'Microsoft.Excel.Objectへの参照設定も削除しておくこと
'Imports Excel = Microsoft.Office.Interop.Excel

'Util.Excelクラスの分割定義
Partial Public Class Excel

    ''' <summary>
    ''' Excel.Rangeオブジェクト保持クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Range
        Implements IDisposable

        Private ReadOnly _range As Object

        Friend ReadOnly Property Value() As Object
            Get
                Return Me._range
            End Get
        End Property

        Friend Sub New(ByRef rng As Object)
            Me._range = rng
        End Sub


        Private _disposedValue As Boolean ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me._disposedValue Then
                If disposing Then
                    If (Me._range IsNot Nothing) Then
                        Runtime.InteropServices.Marshal.ReleaseComObject(Me._range)
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

End Class
