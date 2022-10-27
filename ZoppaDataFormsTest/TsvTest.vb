Imports System
Imports System.Text
Imports Xunit
Imports ZoppaDataForms
Imports ZoppaDataForms.Csv
Imports ZoppaDataForms.Tsv

Public Class TsvTest

    Public Sub New()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    End Sub

    <Fact>
    Sub SplitTest()
        Dim tsv1 = TsvSpliter.CreateSpliter("あ" & vbTab & "い" & vbTab & "う" & vbTab & "え" & vbTab & """お" & vbTab & "を""").Split()
        Assert.Equal(tsv1(0).UnEscape(), "あ")
        Assert.Equal(tsv1(1).UnEscape(), "い")
        Assert.Equal(tsv1(2).UnEscape(), "う")
        Assert.Equal(tsv1(3).UnEscape(), "え")
        Assert.Equal(tsv1(4).UnEscape(), "お" & vbTab & "を")
        Assert.Equal(tsv1(4).Text, """お" & vbTab & "を""")

        Dim tsv2 = TsvSpliter.CreateSpliter("あ" & vbTab & "い" & vbTab & "う" & vbTab & "え" & vbTab & " """"お" & vbTab & "を""""").Split()
        Assert.Equal(tsv2(0).UnEscape(), "あ")
        Assert.Equal(tsv2(1).UnEscape(), "い")
        Assert.Equal(tsv2(2).UnEscape(), "う")
        Assert.Equal(tsv2(3).UnEscape(), "え")
        Assert.Equal(tsv2(4).UnEscape(), "お")
        Assert.Equal(tsv2(5).UnEscape(), "を")

        Dim tsv3 = TsvSpliter.CreateSpliter("あ" & vbTab & "い" & vbTab & "う" & vbTab & "え" & vbTab & """""お" & vbTab & "を""""").Split()
        Assert.Equal(tsv3(0).UnEscape(), "あ")
        Assert.Equal(tsv3(1).UnEscape(), "い")
        Assert.Equal(tsv3(2).UnEscape(), "う")
        Assert.Equal(tsv3(3).UnEscape(), "え")
        Assert.Equal(tsv3(4).UnEscape(), "お")
        Assert.Equal(tsv3(5).UnEscape(), "を")

        Dim tsv4 = TsvSpliter.CreateSpliter("あ" & vbTab & "い" & vbTab & "う" & vbTab & "え" & vbTab & " ""お" & vbTab & "を""").Split()
        Assert.Equal(tsv4(0).UnEscape(), "あ")
        Assert.Equal(tsv4(1).UnEscape(), "い")
        Assert.Equal(tsv4(2).UnEscape(), "う")
        Assert.Equal(tsv4(3).UnEscape(), "え")
        Assert.Equal(tsv4(4).UnEscape(), "お" & vbTab & "を")

        Dim tsv5 = TsvSpliter.CreateSpliter("あ" & vbTab & "い" & vbTab & "う" & vbTab & "え" & vbTab & " ""お"" ""お" & vbTab & "を""").Split()
        Assert.Equal(tsv5(0).UnEscape(), "あ")
        Assert.Equal(tsv5(1).UnEscape(), "い")
        Assert.Equal(tsv5(2).UnEscape(), "う")
        Assert.Equal(tsv5(3).UnEscape(), "え")
        Assert.Equal(tsv5(4).UnEscape(), "お お" & vbTab & "を")

        Dim tsv6 = TsvSpliter.CreateSpliter("あ" & vbTab & "い" & vbTab & "う" & vbTab & "え" & vbTab & """お"" ""お" & vbTab & "を""").Split()
        Assert.Equal(tsv6(0).UnEscape(), "あ")
        Assert.Equal(tsv6(1).UnEscape(), "い")
        Assert.Equal(tsv6(2).UnEscape(), "う")
        Assert.Equal(tsv6(3).UnEscape(), "え")
        Assert.Equal(tsv6(4).UnEscape(), "お お" & vbTab & "を")

        Dim ans As New List(Of Sample1Tsv)()
        Using sr As New TsvStreamReader("TsvFiles\Sample1.tsv", Encoding.GetEncoding("shift_jis"))
            ans = sr.Where(Of Sample1Tsv)(
                    Function(row, item) row >= 1,
                    Function(row, item) New Sample1Tsv(item(0).UnEscape(), item(1).UnEscape(), item(2).UnEscape())
                ).ToList()
        End Using
        Assert.Equal(4, ans.Count)
        Assert.Equal(New Sample1Tsv("1", "2", "3"), ans(0))
        Assert.Equal(New Sample1Tsv($"1{vbTab}1", $"2{vbCrLf}2", $" 3{vbTab}""3 "), ans(1))
        Assert.Equal(New Sample1Tsv("a", "b", "c"), ans(2))
        Assert.Equal(New Sample1Tsv("1", "2", "3"), ans(3))
    End Sub

End Class

Class Sample1Tsv

    Public ReadOnly Property Item1 As String

    Public ReadOnly Property Item2 As String

    Public ReadOnly Property Item3 As String

    Public Sub New(s1 As String, s2 As String, s3 As String)
        Me.Item1 = s1
        Me.Item2 = s2
        Me.Item3 = s3
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, Sample1Tsv)
        Return Me.Item1 = other.Item1 AndAlso
                   Me.Item2 = other.Item2 AndAlso
                   Me.Item3 = other.Item3
    End Function

End Class
