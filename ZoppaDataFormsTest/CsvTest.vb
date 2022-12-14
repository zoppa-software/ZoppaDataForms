Imports System
Imports System.Text
Imports Xunit
Imports ZoppaDataForms
Imports ZoppaDataForms.Csv

Public Class CsvTest

    Public Sub New()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    End Sub

    <Fact>
    Sub ExcelSplitTest()
        Dim csv1 = ExcelCsvSpliter.CreateSpliter("あ,い,う,え,""お,を""").Split()
        Assert.Equal(csv1(0).UnEscape(), "あ")
        Assert.Equal(csv1(1).UnEscape(), "い")
        Assert.Equal(csv1(2).UnEscape(), "う")
        Assert.Equal(csv1(3).UnEscape(), "え")
        Assert.Equal(csv1(4).UnEscape(), "お,を")
        Assert.Equal(csv1(4).Text, """お,を""")

        Dim csv2 = ExcelCsvSpliter.CreateSpliter("あ,い,う,え, """"お,を""""").Split()
        Assert.Equal(csv2(0).UnEscape(), "あ")
        Assert.Equal(csv2(1).UnEscape(), "い")
        Assert.Equal(csv2(2).UnEscape(), "う")
        Assert.Equal(csv2(3).UnEscape(), "え")
        Assert.Equal(csv2(4).UnEscape(), " """"お")
        Assert.Equal(csv2(5).UnEscape(), "を""""")

        Dim csv3 = ExcelCsvSpliter.CreateSpliter("あ,い,う,え,""""お,を""""").Split()
        Assert.Equal(csv3(0).UnEscape(), "あ")
        Assert.Equal(csv3(1).UnEscape(), "い")
        Assert.Equal(csv3(2).UnEscape(), "う")
        Assert.Equal(csv3(3).UnEscape(), "え")
        Assert.Equal(csv3(4).UnEscape(), "お")
        Assert.Equal(csv3(5).UnEscape(), "を""""")

        Dim csv4 = ExcelCsvSpliter.CreateSpliter("あ,い,う,え, ""お,を""").Split()
        Assert.Equal(csv4(0).UnEscape(), "あ")
        Assert.Equal(csv4(1).UnEscape(), "い")
        Assert.Equal(csv4(2).UnEscape(), "う")
        Assert.Equal(csv4(3).UnEscape(), "え")
        Assert.Equal(csv4(4).UnEscape(), " ""お")
        Assert.Equal(csv4(5).UnEscape(), "を""")

        Dim csv5 = ExcelCsvSpliter.CreateSpliter("あ,い,う,え, ""お"" ""お,を""").Split()
        Assert.Equal(csv5(0).UnEscape(), "あ")
        Assert.Equal(csv5(1).UnEscape(), "い")
        Assert.Equal(csv5(2).UnEscape(), "う")
        Assert.Equal(csv5(3).UnEscape(), "え")
        Assert.Equal(csv5(4).UnEscape(), " ""お"" ""お")
        Assert.Equal(csv5(5).UnEscape(), "を""")

        Dim csv6 = ExcelCsvSpliter.CreateSpliter("あ,い,う,え,""お"" ""お,を""").Split()
        Assert.Equal(csv6(0).UnEscape(), "あ")
        Assert.Equal(csv6(1).UnEscape(), "い")
        Assert.Equal(csv6(2).UnEscape(), "う")
        Assert.Equal(csv6(3).UnEscape(), "え")
        Assert.Equal(csv6(4).UnEscape(), "お ""お")
        Assert.Equal(csv6(5).UnEscape(), "を""")

        Dim ans As New List(Of Sample1Csv)()
        Using sr As New ExcelCsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
            ans = sr.Where(Of Sample1Csv)(
                    Function(row, item) row >= 1,
                    Function(row, item) New Sample1Csv(item(0).UnEscape(), item(1).UnEscape(), item(2).UnEscape())
                ).ToList()
        End Using
        Assert.Equal(4, ans.Count)
        Assert.Equal(New Sample1Csv("1", "2", "3"), ans(0))
        Assert.Equal(New Sample1Csv("1,1", $"2{vbCrLf}2", " 3.""3 "), ans(1))
        Assert.Equal(New Sample1Csv("a", "b", "c"), ans(2))
        Assert.Equal(New Sample1Csv("1", "2", "3"), ans(3))
    End Sub

    <Fact>
    Sub SplitTest()
        Dim csv1 = CsvSpliter.CreateSpliter("あ,い,う,え,""お,を""").Split()
        Assert.Equal(csv1(0).UnEscape(), "あ")
        Assert.Equal(csv1(1).UnEscape(), "い")
        Assert.Equal(csv1(2).UnEscape(), "う")
        Assert.Equal(csv1(3).UnEscape(), "え")
        Assert.Equal(csv1(4).UnEscape(), "お,を")
        Assert.Equal(csv1(4).Text, """お,を""")

        Dim csv2 = CsvSpliter.CreateSpliter("あ,い,う,え, """"お,を""""").Split()
        Assert.Equal(csv2(0).UnEscape(), "あ")
        Assert.Equal(csv2(1).UnEscape(), "い")
        Assert.Equal(csv2(2).UnEscape(), "う")
        Assert.Equal(csv2(3).UnEscape(), "え")
        Assert.Equal(csv2(4).UnEscape(), "お")
        Assert.Equal(csv2(5).UnEscape(), "を")

        Dim csv3 = CsvSpliter.CreateSpliter("あ,い,う,え,""""お,を""""").Split()
        Assert.Equal(csv3(0).UnEscape(), "あ")
        Assert.Equal(csv3(1).UnEscape(), "い")
        Assert.Equal(csv3(2).UnEscape(), "う")
        Assert.Equal(csv3(3).UnEscape(), "え")
        Assert.Equal(csv3(4).UnEscape(), "お")
        Assert.Equal(csv3(5).UnEscape(), "を")

        Dim csv4 = CsvSpliter.CreateSpliter("あ,い,う,え, ""お,を""").Split()
        Assert.Equal(csv4(0).UnEscape(), "あ")
        Assert.Equal(csv4(1).UnEscape(), "い")
        Assert.Equal(csv4(2).UnEscape(), "う")
        Assert.Equal(csv4(3).UnEscape(), "え")
        Assert.Equal(csv4(4).UnEscape(), "お,を")

        Dim csv5 = CsvSpliter.CreateSpliter("あ,い,う,え, ""お"" ""お,を""").Split()
        Assert.Equal(csv5(0).UnEscape(), "あ")
        Assert.Equal(csv5(1).UnEscape(), "い")
        Assert.Equal(csv5(2).UnEscape(), "う")
        Assert.Equal(csv5(3).UnEscape(), "え")
        Assert.Equal(csv5(4).UnEscape(), "お お,を")

        Dim csv6 = CsvSpliter.CreateSpliter("あ,い,う,え,""お"" ""お,を""").Split()
        Assert.Equal(csv6(0).UnEscape(), "あ")
        Assert.Equal(csv6(1).UnEscape(), "い")
        Assert.Equal(csv6(2).UnEscape(), "う")
        Assert.Equal(csv6(3).UnEscape(), "え")
        Assert.Equal(csv6(4).UnEscape(), "お お,を")

        Dim ans As New List(Of Sample1Csv)()
        Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
            ans = sr.Where(Of Sample1Csv)(
                    Function(row, item) row >= 1,
                    Function(row, item) New Sample1Csv(item(0).UnEscape(), item(1).UnEscape(), item(2).UnEscape())
                ).ToList()
        End Using
        Assert.Equal(4, ans.Count)
        Assert.Equal(New Sample1Csv("1", "2", "3"), ans(0))
        Assert.Equal(New Sample1Csv("1,1", $"2{vbCrLf}2", " 3.""3 "), ans(1))
        Assert.Equal(New Sample1Csv("a", "b", "c"), ans(2))
        Assert.Equal(New Sample1Csv("1", "2", "3"), ans(3))
    End Sub

End Class

Class Sample1Csv

    Public ReadOnly Property Item1 As String

    Public ReadOnly Property Item2 As String

    Public ReadOnly Property Item3 As String

    Public Sub New(s1 As String, s2 As String, s3 As String)
        Me.Item1 = s1
        Me.Item2 = s2
        Me.Item3 = s3
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, Sample1Csv)
        Return Me.Item1 = other.Item1 AndAlso
                   Me.Item2 = other.Item2 AndAlso
                   Me.Item3 = other.Item3
    End Function

End Class


