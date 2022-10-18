Option Strict On
Option Explicit On

''' <summary>変換パラメータです。</summary>
Public Interface ICvParameter

    ''' <summary>変換する型を取得します。</summary>
    ''' <returns>変換する型。</returns>
    ReadOnly Property ConvertType As Type

    ''' <summary>分割項目を変換します。</summary>
    ''' <param name="input">変換する分割項目。</param>
    ''' <returns>変換後の値。</returns>
    Function Convert(input As IValueItem) As Object

End Interface
