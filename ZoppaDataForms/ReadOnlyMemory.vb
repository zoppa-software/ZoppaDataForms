Option Strict On
Option Explicit On

''' <summary>�ǂݍ��ݐ�p�������B</summary>
Public Structure ReadOnlyMemory

    ''' <summary>�Q�Ƃ��镶���z��B</summary>
    Private ReadOnly mItems As Char()

    ''' <summary>�J�n�ʒu�B</summary>
    Private ReadOnly mStart As Integer

    ''' <summary>�������B</summary>
    Private ReadOnly mLength As Integer

    ''' <summary>���������擾���܂��B</summary>
    ''' <returns>�������B</returns>
    Public ReadOnly Property Length() As Integer
        Get
            Return Me.mLength
        End Get
    End Property

    ''' <summary>�R���X�g���N�^�B</summary>
    ''' <param name="item">�Q�ƕ����z��B</param>
    Public Sub New(item As Char())
        Me.mItems = item
        Me.mStart = 0
        Me.mLength = item.Length
    End Sub

    ''' <summary>�R���X�g���N�^�B</summary>
    ''' <param name="source">���ɂ����ǂݍ��ݐ�p�������B</param>
    ''' <param name="start">�V�����J�n�ʒu�B</param>
    ''' <param name="length">�V�����������B</param>
    Public Sub New(source As ReadOnlyMemory, start As Integer, length As Integer)
        Me.mItems = source.mItems
        Me.mStart = If(source.mStart <= start, start, source.mStart)
        Dim maxlen = source.mStart + source.mLength
        Me.mLength = If(maxlen >= Me.mStart + length, length, maxlen - Me.mStart)
    End Sub

    ''' <summary>�ǂݍ��ݐ�p���������畔�����擾���܂��B</summary>
    ''' <param name="start">�����J�n�ʒu�B</param>
    ''' <param name="length">�����������B</param>
    ''' <returns>�����̓ǂݍ��ݐ�p�������B</returns>
    Public Function Slice(start As Integer, length As Integer) As ReadOnlyMemory
        Dim ln = If(Me.mLength - start >= length, length, Me.mLength - start)
        Return New ReadOnlyMemory(Me, Me.mStart + start, ln)
    End Function

    ''' <summary>������\�����擾���܂��B</summary>
    ''' <returns>������\���B</returns>
    Public Overrides Function ToString() As String
        Return If(Me.mItems IsNot Nothing, New String(Me.mItems, Me.mStart, Me.mLength), "")
    End Function

End Structure