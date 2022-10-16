Option Strict On
Option Explicit On

''' <summary>項目情報。</summary>
Public Interface ISplitItem

    ''' <summary>エスケープを解除した文字列を返す。</summary>
    ''' <returns><エスケープを解除した文字列。/returns>
    ReadOnly Property UnEscape As String

    ''' <summary>生の文字配列を取得します。</summary>
    ''' <returns>生の文字配列。</returns>
    ReadOnly Property Raw As ReadOnlyMemory

    ''' <summary>項目の文字列を取得する。</summary>
    ''' <returns>項目の文字列。</returns>
    ReadOnly Property Text As String

    ''' <summary>読み込み専用メモリを設定します。</summary>
    ''' <param name="rom">読み込み専用メモリ。</param>
    Sub SetRaw(rom As ReadOnlyMemory)

End Interface
