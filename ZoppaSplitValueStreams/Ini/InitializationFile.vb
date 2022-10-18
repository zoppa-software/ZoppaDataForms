Option Strict On
Option Explicit On

Imports System.Text

''' <summary>INIファイル操作機能です。</summary>
Public NotInheritable Class InitializationFile

    Private mKeyAndValue As New Dictionary(Of Section, Dictionary(Of String, String()))()

    Private Shared mConverters As New Lazy(Of Dictionary(Of Type, ICvParameter))(
        Function()
            Return New Dictionary(Of Type, ICvParameter)()
        End Function
    )

    Public Sub New(lines As List(Of String))
        Dim currentSection As New Section()
        For Each srcln In lines
            Dim ln = TrimComment(srcln)

            Dim s = ln.Trim()
            Dim kv = New String() {"", "", "", ""}
            If s <> "" Then
                If IsSection(s) Then
                    currentSection = New Section(s.Substring(1, s.Length - 2))
                ElseIf IsKeyPair(s, kv) Then
                    If Not Me.mKeyAndValue.ContainsKey(currentSection) Then
                        Me.mKeyAndValue.Add(currentSection, New Dictionary(Of String, String())())
                    End If
                    If Not Me.mKeyAndValue(currentSection).ContainsKey(kv(0)) Then
                        Me.mKeyAndValue(currentSection).Add(kv(0), New String() {kv(2), kv(3)})
                    Else
                        Throw New ArgumentException($"既に同じキーが登録されています。[{currentSection.Name}]{kv(0)}")
                    End If
                Else
                    Throw New ArgumentException("セクション、キー／値以外の値が与えられました")
                End If
            End If
        Next
    End Sub

    Public Shared Function LoadIni(iniString As String) As InitializationFile
        Dim lines As New List(Of String)()
        Using sr As New IO.StringReader(iniString)
            Do While sr.Peek() <> -1
                Dim ln = sr.ReadLine()
                If ln.Trim() <> "" Then
                    lines.Add(ln)
                End If
            Loop
        End Using

        Return New InitializationFile(lines)
    End Function

    Public Shared Function Load(iniFilePath As String, Optional encode As Text.Encoding = Nothing) As InitializationFile
        Dim enc = encode
        If enc Is Nothing Then
            enc = Text.Encoding.Default
        End If

        Dim fi As New IO.FileInfo(iniFilePath)
        If fi.Exists Then
            Dim lines As New List(Of String)()
            Using sr As New IO.StreamReader(iniFilePath, enc)
                Do While sr.Peek() <> -1
                    Dim ln = sr.ReadLine()
                    If ln.Trim() <> "" Then
                        lines.Add(ln)
                    End If
                Loop
            End Using

            Return New InitializationFile(lines)
        Else
            Throw New IO.FileNotFoundException("指定したファイルが存在しません")
        End If
    End Function

    Private Shared Function TrimComment(srcln As String) As String
        Dim esc As Boolean = False
        For i As Integer = 0 To srcln.Length - 1
            If srcln(i) = "\"c AndAlso (i = srcln.Length - 1 OrElse srcln(i + 1) <> "\"c) Then
                esc = True
            ElseIf Not esc AndAlso srcln(i) = ";"c Then
                Return srcln.Substring(0, i)
            ElseIf esc Then
                esc = False
            End If
        Next
        Return srcln
    End Function

    Private Shared Function IsSection(ln As String) As Boolean
        Return (ln.Length > 2 AndAlso ln(0) = "["c AndAlso ln(ln.Length - 1) = "]"c)
    End Function

    Private Shared Function IsUnicode(str As String) As Boolean
        If str(0) = "x"c OrElse str(0) = "X"c Then
            For i As Integer = 1 To 4
                If (str(i) >= "0"c AndAlso str(i) <= "9"c) OrElse
                   (str(i) >= "a"c AndAlso str(i) <= "f"c) OrElse
                   (str(i) >= "A"c AndAlso str(i) <= "F"c) Then
                Else
                    Return False
                End If
            Next
            Return True
        Else
            Return False
        End If
    End Function

    Private Shared Function IsKeyPair(ln As String, outKeyAndValue As String()) As Boolean
        Dim strs = New Text.StringBuilder() {
            New Text.StringBuilder(),
            New Text.StringBuilder()
        }
        Dim index As Integer = 0
        Dim esc As Boolean = False
        Dim endc As Char
        For i As Integer = 0 To ln.Length - 1
            If Not esc Then
                If ln(i) = "\"c AndAlso i < ln.Length - 5 AndAlso IsUnicode(ln.Substring(i + 1, 5)) Then
                    strs(index).Append(ln.Substring(i, 6))
                    i += 5
                ElseIf ln(i) = "\"c AndAlso i < ln.Length - 1 Then
                    strs(index).Append(ln.Substring(i, 2))
                    i += 1
                ElseIf ln(i) = "="c Then
                    index += 1
                    If index > 1 Then
                        Return False
                    End If
                ElseIf ln(i) = """"c OrElse ln(i) = "'"c Then
                    If strs(index).ToString().Trim() = "" Then
                        esc = True
                        endc = ln(i)
                    End If
                    strs(index).Append(ln(i))
                Else
                    strs(index).Append(ln(i))
                End If
            Else
                If (ln(i) = "\"c OrElse ln(i) = endc) AndAlso (i < ln.Length - 1 AndAlso ln(i + 1) = endc) Then
                    strs(index).Append(ln(i + 1))
                    i += 1
                ElseIf ln(i) = endc Then
                    esc = False
                End If
                strs(index).Append(ln(i))
            End If
        Next

        For p As Integer = 0 To 1
            Dim str = strs(p).ToString().Trim()
            Dim unstr As New StringBuilder()

            esc = (str(0) = """"c OrElse str(0) = "'"c)
            For i As Integer = If(esc, 1, 0) To str.Length - 1
                If Not esc Then
                    If str(i) = "\"c AndAlso i < str.Length - 1 Then
                        Dim c As Char? = Nothing
                        Select Case str(i + 1)
                            Case "\"c
                                c = "\"c
                            Case "0"c
                                c = CChar(vbNullChar)
                            Case "t"c, "T"c
                                c = CChar(vbTab)
                            Case "r"c, "R"c
                                c = CChar(vbCr)
                            Case "n"c, "N"c
                                c = CChar(vbLf)
                            Case ";"c, "#"c, "="c, ":"c, "\"c
                                c = str(i + 1)
                            Case "x"c, "X"c
                                If i < str.Length - 5 Then
                                    Try
                                        Dim num As Integer = 0
                                        For j As Integer = i + 2 To i + 5
                                            num = (num << 4) + Convert.ToInt32($"{str(j)}", 16)
                                        Next
                                        c = ChrW(num)
                                        i += 4
                                    Catch ex As Exception
                                        ' 空実装
                                    End Try
                                End If
                        End Select

                        If c.HasValue Then
                            unstr.Append(c)
                            i += 1
                        End If
                    Else
                        unstr.Append(str(i))
                    End If
                Else
                    If (str(i) = "\"c OrElse str(i) = endc) AndAlso (i < str.Length - 1 AndAlso str(i + 1) = endc) Then
                        unstr.Append(str(i + 1))
                        i += 1
                    ElseIf str(i) = endc Then
                        esc = False
                    Else
                        unstr.Append(str(i))
                    End If
                End If
            Next

            outKeyAndValue(p * 2) = unstr.ToString()
            outKeyAndValue(p * 2 + 1) = str
        Next
        Return (index = 1)
    End Function

    Public Function GetNoSecssionValue(key As String, Optional defaultValue As String = "") As ValueResult
        Return GetValue(Nothing, key, defaultValue)
    End Function

    Public Function GetValue(secssion As String, key As String, Optional defaultValue As String = "") As ValueResult
        Dim sec = If(secssion Is Nothing, New Section(), New Section(secssion))
        Dim val As String() = Nothing
        If Me.mKeyAndValue.ContainsKey(sec) AndAlso Me.mKeyAndValue(sec).TryGetValue(key, val) Then
            Return New ValueResult(True, val(0), val(1))
        Else
            Return New ValueResult(False, defaultValue, defaultValue)
        End If
    End Function

    Private NotInheritable Class Section

        Public ReadOnly Property DefaultSection As Boolean

        Public ReadOnly Property Name As String

        Public Sub New()
            Me.DefaultSection = True
            Me.Name = ""
        End Sub

        Public Sub New(secname As String)
            Me.DefaultSection = False
            Me.Name = secname
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is Section Then
                Dim other = CType(obj, Section)
                Return (Me.DefaultSection = other.DefaultSection AndAlso Me.Name = other.Name)
            Else
                Return False
            End If
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Me.DefaultSection.GetHashCode() Xor Me.Name.GetHashCode()
        End Function

    End Class

    Public Structure ValueResult
        Implements IValueItem

        Public ReadOnly Property IsSome As Boolean

        Public ReadOnly Property UnEscape As String Implements IValueItem.UnEscape

        Public ReadOnly Property Text As String Implements IValueItem.Text

        Public Sub New(sm As Boolean, unEsc As String, text As String)
            Me.IsSome = sm
            Me.UnEscape = unEsc
            Me.Text = text
        End Sub

        Public Function Convert(Of TResult, TConverter As {ICvParameter, New})() As TResult
            Dim converters = mConverters.Value
            Dim cvKey = GetType(TConverter)
            If Not converters.ContainsKey(cvKey) Then
                converters.Add(cvKey, New TConverter())
            End If

            If Me.IsSome Then
                Try
                    Return CType(converters(cvKey).Convert(Me), TResult)
                Catch ex As Exception
                    Throw New InvalidCastException("値の型変換に失敗しました")
                End Try
            Else
                Return Nothing
            End If
        End Function

    End Structure

End Class
