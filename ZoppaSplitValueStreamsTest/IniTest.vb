Imports System
Imports System.Text
Imports Xunit
Imports ZoppaSplitValueStreams

Public Class IniTest

    Public Sub New()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    End Sub

    <Fact>
    Sub ReadTest()
        Dim iniFile = InitializationFile.Load("IniFiles\Sample1.ini", Encoding.GetEncoding("shift_jis"))
        Dim a1 = iniFile.GetNoSecssionValue("KEY1")
        Assert.True(a1.IsSome)
        Assert.Equal("     ", a1.UnEscape)

        Dim a2 = iniFile.GetNoSecssionValue("KEY2", "XXX")
        Assert.False(a2.IsSome)
        Assert.Equal("XXX", a2.UnEscape)

        Dim a3 = iniFile.GetValue("SECTION1", "KEY1")
        Assert.True(a3.IsSome)
        Assert.Equal("\keydata1", a3.UnEscape)

        Dim a4 = iniFile.GetValue("SECTION1", "KEY2")
        Assert.True(a4.IsSome)
        Assert.Equal("key=data2", a4.UnEscape)

        Dim a5 = iniFile.GetValue("SECTION2", "KEYA")
        Assert.True(a5.IsSome)
        Assert.Equal("keydataA", a5.UnEscape)

        Dim a6 = iniFile.GetValue("SPECIAL", "KEYZ")
        Assert.True(a6.IsSome)
        Assert.Equal(";#=:烏" & vbCrLf & "改行テスト" & vbNullChar, a6.UnEscape)
    End Sub

    <Fact>
    Sub EscapeTest()
        Dim iniFile = InitializationFile.Load("IniFiles\Sample2.ini", Encoding.GetEncoding("shift_jis"))
        Dim a1 = iniFile.GetNoSecssionValue("KEY1")
        Assert.Equal("c:\user\desktop", a1.UnEscape)

        Dim a2 = iniFile.GetNoSecssionValue("KEY2")
        Assert.Equal("c:\user\desktop", a2.UnEscape)

        Dim a3 = iniFile.GetNoSecssionValue("KEY3")
        Assert.Equal("123""456", a3.UnEscape)

        Dim a4 = iniFile.GetNoSecssionValue("KEY4")
        Assert.Equal("123""456", a4.UnEscape)
    End Sub

End Class
