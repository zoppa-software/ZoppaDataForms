Option Strict On
Option Explicit On

''' <summary>パラメータコンバータ定義モジュールです。</summary>
Public Module Parameter

    ''' <summary>バイト配列パラメータコンバータです。</summary>
    Public ReadOnly CvBinary As ICvParameter = New BinaryConverter()

    ''' <summary>バイトパラメータコンバータです。</summary>
    Public ReadOnly CvByte As ICvParameter = New ByteConverter()

    ''' <summary>真偽値パラメータコンバータです。</summary>
    Public ReadOnly CvBoolean As ICvParameter = New BooleanConverter()

    ''' <summary>日付型パラメータコンバータです。</summary>
    Public ReadOnly CvDateTime As ICvParameter = New DateConverter()

    ''' <summary>Decimalパラメータコンバータです。</summary>
    Public ReadOnly CvDecimal As ICvParameter = New DecimalConverter()

    ''' <summary>Doubleパラメータコンバータです。</summary>
    Public ReadOnly CvDouble As ICvParameter = New DoubleConverter()

    ''' <summary>Shortパラメータコンバータです。</summary>
    Public ReadOnly CvShort As ICvParameter = New ShortConverter()

    ''' <summary>Integerパラメータコンバータです。</summary>
    Public ReadOnly CvInteger As ICvParameter = New IntegerConverter()

    ''' <summary>Longパラメータコンバータです。</summary>
    Public ReadOnly CvLong As ICvParameter = New LongConverter()

    ''' <summary>Singleパラメータコンバータです。</summary>
    Public ReadOnly CvSingle As ICvParameter = New SingleConverter()

    ''' <summary>文字列パラメータコンバータです。</summary>
    Public ReadOnly CvString As ICvParameter = New StringConverter()

    ''' <summary>時間型パラメータコンバータです。</summary>
    Public ReadOnly CvTime As ICvParameter = New TimeConverter()

    ''' <summary>バイト配列に変換します。</summary>
    Private NotInheritable Class BinaryConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Byte())
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Dim ans As New List(Of Byte)()

            Dim buf = New Char(1) {}
            Using sr As New IO.StringReader(If(inp.Length Mod 2 = 0, inp, "0" & inp))
                Do While sr.Peek() <> -1
                    sr.Read(buf, 0, 2)

                    Dim b = 0
                    For i As Integer = 0 To 1
                        Select Case buf(i)
                            Case "0"c To "9"c
                                b = b * 16 + (AscW(buf(i)) - AscW("0"c))
                            Case "a"c To "f"c
                                b = b * 16 + 10 + (AscW(buf(i)) - AscW("a"c))
                            Case "A"c To "F"c
                                b = b * 16 + 10 + (AscW(buf(i)) - AscW("A"c))
                        End Select
                    Next
                    ans.Add(CByte(b))
                Loop
            End Using

            Return ans.ToArray()
        End Function
    End Class

    ''' <summary>バイトに変換します。</summary>
    Private NotInheritable Class ByteConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Byte?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToByte(inp), Nothing)
        End Function
    End Class

    ''' <summary>真偽値に変換します。</summary>
    Private NotInheritable Class BooleanConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Boolean?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToBoolean(inp), Nothing)
        End Function
    End Class

    ''' <summary>日付型に変換します。</summary>
    Private NotInheritable Class DateConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Date?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToDateTime(inp), Nothing)
        End Function
    End Class

    ''' <summary>Decimal値に変換します。</summary>
    Private NotInheritable Class DecimalConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Decimal?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToDecimal(inp), Nothing)
        End Function
    End Class

    ''' <summary>Double値に変換します。</summary>
    Private NotInheritable Class DoubleConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Double?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToDouble(inp), Nothing)
        End Function
    End Class

    ''' <summary>Short値に変換します。</summary>
    Private NotInheritable Class ShortConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Short?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToInt16(inp), Nothing)
        End Function
    End Class

    ''' <summary>Integer値に変換します。</summary>
    Private NotInheritable Class IntegerConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Integer?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToInt32(inp), Nothing)
        End Function
    End Class

    ''' <summary>Long値に変換します。</summary>
    Private NotInheritable Class LongConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Long?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToInt64(inp), Nothing)
        End Function
    End Class

    ''' <summary>Single値に変換します。</summary>
    Private NotInheritable Class SingleConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(Single?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", System.Convert.ToInt64(inp), Nothing)
        End Function
    End Class

    ''' <summary>文字列に変換します。</summary>
    Private NotInheritable Class StringConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(String)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Return input.UnEscape
        End Function
    End Class

    ''' <summary>時間型に変換します。</summary>
    Private NotInheritable Class TimeConverter
        Implements ICvParameter

        ''' <summary>変換する型を取得します。</summary>
        ''' <returns>変換する型。</returns>
        Public ReadOnly Property ConvertType As Type Implements ICvParameter.ConvertType
            Get
                Return GetType(TimeSpan?)
            End Get
        End Property

        ''' <summary>分割項目を変換します。</summary>
        ''' <param name="input">変換する分割項目。</param>
        ''' <returns>変換後の値。</returns>
        Public Function Convert(input As IValueItem) As Object Implements ICvParameter.Convert
            Dim inp = input.UnEscape
            Return If(inp <> "", TimeSpan.Parse(inp), Nothing)
        End Function
    End Class

End Module
