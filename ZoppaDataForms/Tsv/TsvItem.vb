Option Strict On
Option Explicit On

Imports System.Text

Namespace Tsv

    ''' <summary>CSV項目情報。</summary>
    Public Structure TsvItem
        Implements ISplitItem

        ' 読み込み専用メモリ
        Private mRom As ReadOnlyMemory

        ''' <summary>エスケープを解除した文字列を返す。</summary>
        ''' <returns>エスケープを解除した文字列。</returns>
        Public ReadOnly Property UnEscape As String Implements IValueItem.UnEscape
            Get
                Dim str = Me.Text.Trim()

                Dim buf As New StringBuilder(str.Length)
                Dim esc As Boolean = False
                For i As Integer = 0 To str.Length - 1
                    Dim c = str(i)

                    If c = """"c Then
                        If esc Then
                            If i < str.Length - 1 AndAlso str(i + 1) = """"c Then
                                buf.Append(""""c)
                                i += 1
                            Else
                                esc = False
                            End If
                        Else
                            esc = True
                        End If
                    Else
                        buf.Append(c)
                    End If
                Next
                Return buf.ToString()
            End Get
        End Property

        ''' <summary>生の文字配列を取得します。</summary>
        ''' <returns>生の文字配列。</returns>
        Public ReadOnly Property Raw As ReadOnlyMemory Implements ISplitItem.Raw
            Get
                Return Me.mRom
            End Get
        End Property

        ''' <summary>項目の文字列を取得する。</summary>
        ''' <returns>項目の文字列。</returns>
        Public ReadOnly Property Text As String Implements ISplitItem.Text
            Get
                Return Me.Raw.ToString()
            End Get
        End Property

        ''' <summary>読み込み専用メモリを設定します。</summary>
        ''' <param name="rom">読み込み専用メモリ。</param>
        Public Sub SetRaw(rom As ReadOnlyMemory) Implements ISplitItem.SetRaw
            Me.mRom = rom
        End Sub

    End Structure

End Namespace
