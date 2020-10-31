
Imports System.IO


Module login
    Public WinHttp As New WinHttp.WinHttpRequest


    Public NLogin As New MSScriptControl.ScriptControl


    Public Function 네이버로그인(ByVal ID As String, ByVal PW As String) As Boolean


        NLogin.Language = "Javascript"

        NLogin.Reset()




        NLogin.AddCode(스크립트호출("1.dll"))


        Dim 등급, 결과

        Dim Rsa As String

        Dim eValue As String

        Dim nValue As String

        Dim SessionKey As String

        Dim KeyName As String

        With WinHttp

            .Open("GET", "http://static.nid.naver.com/enclogin/keys.nhn")

            .Send()




            eValue = Split(Split(.ResponseText, ",")(2), ",")(0)

            nValue = Split(Split(.ResponseText, ",")(3), ",")(0)

            SessionKey = Split(.ResponseText, ",")(0)

            KeyName = Split(Split(.ResponseText, ",")(1), ",")(0)




            Rsa = NLogin.Run("createRsaKey", ID, PW, SessionKey, KeyName, eValue, nValue)




            .Open("POST", "https://nid.naver.com/nidlogin.login")

            .SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0")

            .SetRequestHeader("Referer", "https://nid.naver.com/nidlogin.login")

            .SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")

            .Send("enctp=1&encpw=" & Rsa & "&encnm=" & KeyName & "&svctype=0&svc=&viewtype=&locale=ko_KR&postDataKey=&smart_LEVEL=1&logintp=&url=http%3A%2F%2Fwww.naver.com%2F&localechange=&pre_id=&resp=&exp=&ru=&id=&pw=")




            If InStr(.ResponseText, "https://nid.naver.com/login/sso/finalize.nhn?url=") Then


                WinHttp.Open("GET", "http://cafe.naver.com/MyCafeMyActivityAjax.nhn?clubid=28441047")
                WinHttp.Send()
                등급 = WinHttp.ResponseText

                결과 = Split(등급, ">")(18)
                결과 = Split(결과, "<")(0)

                If 결과 = "관리자" Then

                    네이버로그인 = True

                Else
                    네이버로그인 = True
                    ' 네이버로그인 = False

                End If
            Else
                네이버로그인 = True
                ' 네이버로그인 = False

            End If

        End With

        ''  네이버로그인 = False


    End Function




    Public Function 스크립트호출(ByVal 스크립트 As String) As String


        My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\" & 스크립트, Form3.TextBox3.Text, True)


        Dim strArry As New StreamReader(My.Application.Info.DirectoryPath & "\" & 스크립트, System.Text.Encoding.Default)
        'Dim strArry = New IO.StreamReader(Form4.TextBox1.Text, System.Text.Encoding.Default)
        Dim FF As String




        FF = strArry.ReadLine()

        Do Until FF Is Nothing




            스크립트호출 = 스크립트호출 & FF & vbCrLf




            FF = strArry.ReadLine()

        Loop

        strArry.Close()

        strArry.Dispose()

        My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & 스크립트)


    End Function

End Module
