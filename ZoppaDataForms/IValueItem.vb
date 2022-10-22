Option Strict On
Option Explicit On

''' <summary>項目情報。</summary>
Public Interface IValueItem

    ''' <summary>エスケープを解除した文字列を返す。</summary>
    ''' <returns><エスケープを解除した文字列。/returns>
    ReadOnly Property UnEscape As String

    ''' <summary>項目の文字列を取得する。</summary>
    ''' <returns>項目の文字列。</returns>
    ReadOnly Property Text As String

End Interface
