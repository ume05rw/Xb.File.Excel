Option Strict On

'�f�o�b�O�p�L�q�F�[�i���͉��L���g�p���Ȃ��B
'Microsoft.Excel.Object�ւ̎Q�Ɛݒ���폜���Ă�������
'Imports Excel = Microsoft.Office.Interop.Excel

'Util.Excel�N���X�̕�����`
Partial Public Class Excel

    ''' <summary>
    ''' Excel.Range�I�u�W�F�N�g�ێ��N���X
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


        Private _disposedValue As Boolean ' �d������Ăяo�������o����ɂ�

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
        ' ���̃R�[�h�́A�j���\�ȃp�^�[���𐳂��������ł���悤�� Visual Basic �ɂ���Ēǉ�����܂����B
        Public Sub Dispose() Implements IDisposable.Dispose
            ' ���̃R�[�h��ύX���Ȃ��ł��������B�N���[���A�b�v �R�[�h����� Dispose(ByVal disposing As Boolean) �ɋL�q���܂��B
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class

End Class
