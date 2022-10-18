Option Strict On
Option Explicit On
Imports System.Reflection
Imports ZoppaSplitValueStreams.Csv

''' <summary>読み込みストリーム（共通）</summary>
Public MustInherit Class SplitStreamReader(Of TSpliter As {Spliter(Of TItem), New}, TItem As {ISplitItem, New})
    Inherits IO.StreamReader
    Implements IEnumerable(Of Pointer)

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    Public Sub New(stream As IO.Stream)
        MyBase.New(stream)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    Public Sub New(path As String)
        MyBase.New(path)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(stream As IO.Stream, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(stream, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding)
        MyBase.New(stream, encoding)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(path As String, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(path, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    Public Sub New(path As String, encoding As Text.Encoding)
        MyBase.New(path, encoding)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(path, encoding, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
        MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
        MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    ''' <param name="leaveOpen">ストリームを開いたままにするならば真。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, leaveOpen As Boolean)
        MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, leaveOpen)
    End Sub

    ''' <summary>列挙子を取得する。</summary>
    ''' <returns>列挙子。</returns>
    Public Iterator Function GetEnumerator() As IEnumerator(Of Pointer) Implements IEnumerable(Of Pointer).GetEnumerator
        Dim spliter As New TSpliter()
        spliter.SetTextReader(Me)

        Dim ans = spliter.ReadLine()
        Dim index As Integer = 0
        Do While ans.HasResult
            Yield New Pointer(index, ans.Items)
            index += 1
            ans = spliter.ReadLine()
        Loop
    End Function

    ''' <summary>列挙子を取得する。</summary>
    ''' <returns>列挙子。</returns>
    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    ''' <summary>指定した型へ変換して取得するクエリ。</summary>
    ''' <typeparam name="TResult">変換後の型。</typeparam>
    ''' <param name="func">変換するためのラムダ式。</param>
    ''' <returns>変換後型の列挙子。</returns>
    Public Iterator Function SelectCsv(Of TResult)(func As Func(Of Integer, List(Of TItem), TResult)) As IEnumerable(Of TResult)
        For Each item In Me
            Yield func(item.Row, item.Items)
        Next
    End Function

    ''' <summary>指定した型のコンストラクタでインスタンスを生成し取得するクエリ。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="columTypes">コンストラクタの引数。</param>
    ''' <returns>変換後型の列挙子。</returns>
    Public Iterator Function SelectCsv(Of T)(ParamArray columTypes As ICvParameter()) As IEnumerable(Of T)
        ' 引数の配列を作成
        Dim clmTps = columTypes.Select(Function(v) v.ConvertType).ToArray()

        ' コンストラクタを取得する
        Dim constructor = GetConstructor(Of T)(clmTps)

        ' インスタンスを生成しながら返す
        Dim fields As New ArrayList()
        For Each item In Me
            SetFields(columTypes, item, fields)
            Yield CType(constructor.Invoke(fields.ToArray()), T)
        Next
    End Function

    ''' <summary>条件に一致した行を指定した型へ変換して取得するクエリ。</summary>
    ''' <typeparam name="TResult">変換後の型。</typeparam>
    ''' <param name="condition">条件判定するラムダ式。</param>
    ''' <param name="func">変換するためのラムダ式。</param>
    ''' <returns>変換後型の列挙子。</returns>
    Public Iterator Function WhereCsv(Of TResult)(condition As Func(Of Integer, List(Of TItem), Boolean),
                                                  func As Func(Of Integer, List(Of TItem), TResult)) As IEnumerable(Of TResult)
        For Each item In Me
            If condition(item.Row, item.Items) Then
                Yield func(item.Row, item.Items)
            End If
        Next
    End Function

    ''' <summary>条件に一致した行を指定した型のコンストラクタでインスタンスを生成し取得するクエリ。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="condition">条件判定するラムダ式。</param>
    ''' <param name="columTypes">コンストラクタの引数。</param>
    ''' <returns>変換後型の列挙子。</returns>
    Public Iterator Function WhereCsv(Of T)(condition As Func(Of Integer, List(Of TItem), Boolean),
                                            ParamArray columTypes As ICvParameter()) As IEnumerable(Of T)
        ' 引数の配列を作成
        Dim clmTps = columTypes.Select(Function(v) v.ConvertType).ToArray()

        ' コンストラクタを取得する
        Dim constructor = GetConstructor(Of T)(clmTps)

        ' インスタンスを生成しながら返す
        Dim fields As New ArrayList()
        For Each item In Me
            If condition(item.Row, item.Items) Then
                SetFields(columTypes, item, fields)
                Yield CType(constructor.Invoke(fields.ToArray()), T)
            End If
        Next
    End Function

    ''' <summary>コンストラクタの参照を取得します。</summary>
    ''' <typeparam name="TNew">コンストラクタを取得する型。</typeparam>
    ''' <param name="clmTps">コンストラクタの引数。</param>
    ''' <returns>コンストラクタの参照。</returns>
    Private Shared Function GetConstructor(Of TNew)(clmTps() As Type) As ConstructorInfo
        ' コンストラクタの参照を取得
        Dim constructor = GetType(TNew).GetConstructor(clmTps)
        If constructor Is Nothing Then
            constructor = GetType(TNew).GetConstructor(New Type() {GetType(Object())})
        End If

        ' 取得できなければエラー
        If constructor Is Nothing Then
            Dim info = String.Join(",", clmTps.Select(Function(c) c.Name).ToArray())
            Throw New SplitException($"取得データに一致するコンストラクタがありません:{info}")
        End If
        Return constructor
    End Function

    ''' <summary>変換方法に従って項目変換したリストを返す。</summary>
    ''' <param name="columTypes">列の変換方法。</param>
    ''' <param name="item">変換する項目。</param>
    ''' <param name="fields">リスト（戻り値）</param>
    Private Shared Sub SetFields(columTypes As ICvParameter(), item As Pointer, fields As ArrayList)
        fields.Clear()
        For i As Integer = 0 To Math.Min(columTypes.Length, item.Items.Count) - 1
            Try
                fields.Add(columTypes(i).Convert(item.Items(i)))
            Catch ex As Exception
                Throw New SplitException($"変換に失敗しました:{i},{item.Row} {item.Items(i).Raw} -> {columTypes(i).ConvertType.Name}")
            End Try
        Next
    End Sub

    ''' <summary>分割項目リストを取得します。</summary>
    ''' <returns>分割項目リスト。</returns>
    Public Function ReadSplit() As List(Of TItem)
        Dim spliter As New TSpliter()
        spliter.SetTextReader(Me)

        Dim ans = spliter.ReadLine()
        If ans.HasResult Then
            Dim res = New List(Of TItem)(ans.Items.Count)
            For Each i In ans.Items
                Dim item As New TItem()
                item.SetRaw(i)
                res.Add(item)
            Next
            Return res
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>項目のポインタ。</summary>
    Public Structure Pointer

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="row">行位置。</param>
        ''' <param name="items">分割項目。</param>
        Public Sub New(row As Integer, items As List(Of ReadOnlyMemory))
            Me.Row = row
            Me.Items = New List(Of TItem)(items.Count)
            For Each i In items
                Dim item As New TItem()
                item.SetRaw(i)
                Me.Items.Add(item)
            Next
        End Sub

        ''' <summary>行位置を取得する。</summary>
        ''' <returns>行位置。</returns>
        Public ReadOnly Property Row() As Integer

        ''' <summary>分割項目の配列を取得する。</summary>
        ''' <returns>分割項目配列。</returns>
        Public ReadOnly Property Items() As List(Of TItem)

        ''' <summary>文字表現を取得する。</summary>
        ''' <returns>文字列。</returns>
        Public Overrides Function ToString() As String
            Return $"row index:{Me.Row} colum count:{Me.Items.Count}"
        End Function

    End Structure

End Class
