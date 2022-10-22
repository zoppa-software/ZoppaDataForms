Option Strict On
Option Explicit On

Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices

''' <summary>文字列分割機能（共通）</summary>
Public MustInherit Class Spliter(Of T As {ISplitItem, New})

    ''' <summary>内部ストリームを取得します。</summary>
    Private mInnerStream As TextReader

    ''' <summary>コンストラクタ。</summary>
    Public Sub New()

    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="inputStream">入力ストリーム。</param>
    Protected Sub New(inputStream As StreamReader)
        Me.mInnerStream = inputStream
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="inputText">入力文字列。</param>
    Protected Sub New(inputText As String)
        Me.mInnerStream = New StringReader(inputText)
    End Sub

    ''' <summary>テキストストリームを設定します。</summary>
    ''' <param name="reader">テキストストリーム。</param>
    Public Sub SetTextReader(reader As TextReader)
        Me.mInnerStream = reader
    End Sub

    ''' <summary>一行読み込み、分割を行います。</summary>
    ''' <returns>分割結果。</returns>
    Public Function ReadLine() As ReadResult
        If Me.mInnerStream IsNot Nothing Then
            With Me.ReadLine(Me.mInnerStream)
                Return New ReadResult(.readChars, .splitPoint)
            End With
        Else
            Throw New NullReferenceException("テキストストリームが設定されていません")
        End If
    End Function

    ''' <summary>一行読み込み、指定した型のコンストラクタでインスタンスを生成し取得するクエリ。</summary>
    ''' <typeparam name="TResult">変換後の型。</typeparam>
    ''' <param name="columTypes">コンストラクタの引数。</param>
    ''' <returns>生成したインスタンス。</returns>
    Public Function ReadLine(Of TResult)(ParamArray columTypes As ICvParameter()) As T
        ' 引数の配列を作成
        Dim clmTps = columTypes.Select(Function(v) v.ConvertType).ToArray()

        ' コンストラクタを取得する
        Dim constructor = GetConstructor(Of T)(clmTps)

        ' 読み込めたらインスタンスを生成して返す
        Dim item = Me.ReadLine()
        If item.HasResult Then
            Dim fields As New ArrayList()
            SetFields(columTypes, item.Items, fields)
            Return CType(constructor.Invoke(fields.ToArray()), T)
        Else
            Return Nothing
        End If
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
    Private Shared Sub SetFields(columTypes As ICvParameter(), items As List(Of ReadOnlyMemory), fields As ArrayList)
        fields.Clear()
        For i As Integer = 0 To Math.Min(columTypes.Length, items.Count) - 1
            Try
                Dim itm As New T()
                itm.SetRaw(items(i))
                fields.Add(columTypes(i).Convert(itm))
            Catch ex As Exception
                Throw New SplitException($"変換に失敗しました:{i},{items(i)} -> {columTypes(i).ConvertType.Name}")
            End Try
        Next
    End Sub

    ''' <summary>一行読み込み、分割を行います。</summary>
    ''' <param name="readStream">読み込みストリーム。</param>
    ''' <returns>読み込んだ文字列と分割位置リスト。</returns>
    Protected MustOverride Function ReadLine(readStream As TextReader) As (readChars As Char(), splitPoint As List(Of Integer))

    ''' <summary>内部より一行を読み込み、分割して返します。</summary>
    ''' <returns>分割した項目の配列。</returns>
    Public Function Split() As List(Of T)
        Dim ans = Me.ReadLine()
        Dim res As New List(Of T)(ans.Items.Count)
        For Each i In ans.Items
            Dim itm As New T()
            itm.SetRaw(i)
            res.Add(itm)
        Next
        Return res
    End Function

    ''' <summary>分割結果を保持する。</summary>
    Public Structure ReadResult

        ''' <summary>分割位置リスト。</summary>
        Private ReadOnly mSPoint As List(Of Integer)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="chars">読み込んだ文字配列。</param>
        ''' <param name="spoint">分割位置。</param>
        Public Sub New(chars As Char(), spoint As List(Of Integer))
            Me.Chars = chars
            Me.mSPoint = spoint
        End Sub

        ''' <summary>読み込んだ文字配列を取得します。</summary>
        ''' <returns>文字配列。</returns>
        Public ReadOnly Property Chars() As Char()

        ''' <summary>項目があれば真を返します。</summary>
        ''' <returns>項目があれば真。</returns>
        Public ReadOnly Property HasResult() As Boolean
            Get
                Return (Me.Chars.Length > 0)
            End Get
        End Property

        ''' <summary>項目のリストを返します。</summary>
        ''' <returns>項目リスト。</returns>
        Public ReadOnly Property Items As List(Of ReadOnlyMemory)
            Get
                Dim src As New ReadOnlyMemory(Me.Chars)
                Dim split As New List(Of ReadOnlyMemory)(Me.mSPoint.Count - 1)
                For i As Integer = 0 To Me.mSPoint.Count - 2
                    split.Add(src.Slice(Me.mSPoint(i), (Me.mSPoint(i + 1) - 1) - Me.mSPoint(i)))
                Next
                Return split
            End Get
        End Property

    End Structure

End Class
