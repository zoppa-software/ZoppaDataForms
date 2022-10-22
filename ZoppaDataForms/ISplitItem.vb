Option Strict On
Option Explicit On

''' <summary>項目情報。</summary>
Public Interface ISplitItem
    Inherits IValueItem

    ''' <summary>生の文字配列を取得します。</summary>
    ''' <returns>生の文字配列。</returns>
    ReadOnly Property Raw As ReadOnlyMemory

    ''' <summary>読み込み専用メモリを設定します。</summary>
    ''' <param name="rom">読み込み専用メモリ。</param>
    Sub SetRaw(rom As ReadOnlyMemory)

End Interface
