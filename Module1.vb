Imports System
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Security.Cryptography
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Collections.Generic
Imports System.Text
Imports System.IO

Module Module1

    Private ReadOnly UnixEpoch As DateTime = New DateTime(1970, 1, 1, 9, 0, 0, DateTimeKind.Utc)
    Private orderList As ArrayList = New ArrayList()

    Dim FileName As String
    Dim APIID As String
    Dim APIKEY As String

    Sub Main()

        Dim args() As String
        Dim num As Integer

        args = Split(Command(), " ")
        '引数１：ダウンロードファイル名
        '引数２：APIID
        '引数２：APIKEY

        If args(0) = "help" Or args(0) = "HELP" Then
            Console.WriteLine("ZaifTradeHistory.exe ファイル名 APIID APIKEY")
        End If

        ' 引数チェック
        num = args.Count

        If num < 3 Then
            Console.WriteLine("引数が足りません。")
            Exit Sub
        End If

        FileName = args(0)
        APIID = args(1)
        APIKEY = args(2)

        orderList.Clear()

        Dim chk As Int16 = 1
        Dim Debug As Int16 = 0
        Dim number As Int32 = 0

        If Debug = 0 Then
            Do While chk = 1
                chk = getTradeHistory(number)

                If chk = 2 Then
                    chk = 1
                Else
                    number += 1000
                End If

                System.Threading.Thread.Sleep(1000 * 5)
            Loop

        Else
            number = 118000 - 200
            chk = getTradeHistory(number)

        End If

        orderList.Sort()


        Dim tmpData As String
        'Dim FileName As String = "D:\Tax\2018Crypt\Zaif\tradeHistory_Zaif118000-100.csv"

        Dim Fw As IO.StreamWriter

        Fw = New IO.StreamWriter(FileName, False, System.Text.Encoding.Default)

        For Each history As orderHistory In orderList

            tmpData = ConstructLine(history.TradeDate.ToString, "," + history.Pair, "," + history.Order.ToString,
                                             "," + history.Price.ToString, "," + history.Size.ToString, "," + history.Fee.ToString)
            Fw.WriteLine(tmpData)
        Next


        Fw.Close()


    End Sub

    Private Function ConstructLine(ByVal str1 As String, ByVal str2 As String, ByVal str3 As String, ByVal str4 As String,
                           ByVal str5 As String, ByVal str6 As String) As String
        Dim strData As String

        strData = str1

        strData = strData + str2
        strData = strData + str3
        strData = strData + str4
        strData = strData + str5
        strData = strData + str6

        Return strData

    End Function


    Private Function getTradeHistory(ByVal iFrom As Int32) As Int16

        Dim endpointUri As Uri = New Uri("https://api.zaif.jp")
        Dim method As String = "GET"
        Dim method1 As String = ""
        Dim path As String = ""
        Dim query As String = ""

        Dim jName As String
        Dim tmpData As String
        Dim err As Int16 = 0

        method = "POST"
        method1 = "trade_history"
        path = "tapi"
        query = ""

        Dim TimeNow As Long = (Date.UtcNow.Ticks - DateTime.Parse("1970-01-01 00:00:00").Ticks) / 10000000
        Dim timestamp As String = TimeNow.ToString

        Dim parameters As Dictionary(Of String, String) = New Dictionary(Of String, String)
        parameters.Add("from", iFrom)
        parameters.Add("count", 1000)

        parameters.Add("nonce", timestamp)
        parameters.Add("method", method1)

        Dim content As FormUrlEncodedContent = New FormUrlEncodedContent(parameters)
        Dim task1 = Task.Run(Async Function()
                                 query = Await content.ReadAsStringAsync()
                                 content = New FormUrlEncodedContent(parameters)
                             End Function)
        task1.Wait()

        Dim request As HttpRequestMessage = New HttpRequestMessage(New HttpMethod(method), path)
        Dim hash As String = CreateDigitalSignature(query, APIKEY)
        request.Headers.Add("Key", APIID)
        request.Headers.Add("Sign", hash)
        request.Content = content

        Dim task2 = Task.Run(Async Function()
                                 'Await _semaphore.WaitAsync()
                                 Dim client As HttpClient = New HttpClient()

                                 client.BaseAddress = endpointUri

                                 Dim message As HttpResponseMessage = Await (client.SendAsync(request))
                                 Dim response_json As String = Await (message.Content.ReadAsStringAsync())

                                 Try
                                     Dim JsonObject As Object = JsonConvert.DeserializeObject(response_json)

                                     If JsonObject("success") = 0 Then
                                         err = 1
                                     End If

                                     Dim JsonObject1 As Object = JsonObject("return")
                                     Dim jTokens As JEnumerable(Of JToken) = JsonObject1.Children

                                     If jTokens.Count = 0 Then
                                         err = 2
                                         Exit Try
                                     End If

                                     For Each jT As JToken In jTokens

                                         If jT.Type = JTokenType.Property Then
                                             Dim jP As JProperty = jT
                                             jName = jP.Name

                                             Dim tempJsonObject As Object = JsonObject1(jName)
                                             Dim tHistory As orderHistory = New orderHistory

                                             tHistory.Pair = CType(tempJsonObject("currency_pair"), String)
                                             tHistory.Price = CType(tempJsonObject("price"), Single)
                                             tHistory.Size = CType(tempJsonObject("amount"), Single)
                                             tHistory.Fee = CType(tempJsonObject("fee"), Single)
                                             tHistory.Bonus = CType(tempJsonObject("bonus"), Single)

                                             If CType(tempJsonObject("your_action"), String) = "ask" Then
                                                 tHistory.Order = "sell"
                                             ElseIf CType(tempJsonObject("your_action"), String) = "bid" Then
                                                 tHistory.Order = "buy"
                                             Else
                                                 tHistory.Order = "err"
                                             End If

                                             tmpData = CType(tempJsonObject("timestamp"), String)
                                             tHistory.TradeDate = FromUnixTime(CType(tmpData, Long))

                                             orderList.Add(tHistory)

                                         End If

                                     Next


                                 Catch ex As Exception
                                     Throw ex
                                 End Try


                             End Function)

        Try
            task2.Wait()
        Catch ex As Exception
            'Throw ex
            Console.WriteLine("リトライNumber = " + iFrom.ToString)
            err = 3
            Exit Try
        End Try

        If err = 1 Or err = 2 Then
            Return 0
        ElseIf err = 3 Then
            Return 2
        End If

        Return 1

    End Function

    Private Function CreateDigitalSignature(ByVal message As String,
                                         ByVal privateKey As String) As String

        Dim encoding As New System.Text.ASCIIEncoding
        Dim key() As Byte = encoding.GetBytes(privateKey)
        Dim XML() As Byte = encoding.GetBytes(message)
        Dim myHMACSHA512 As New System.Security.Cryptography.HMACSHA512(key)
        Dim HashCode As Byte() = myHMACSHA512.ComputeHash(XML)
        Dim output As String = ""

        For i = 0 To HashCode.Count - 1
            output += HashCode(i).ToString("x2")
        Next

        Return output


    End Function

    Public Function FromUnixTime(ByVal unixTime As Long) As DateTime

        ' unix epochからunixTime秒だけ経過した時刻を求める
        Return UnixEpoch.AddSeconds(unixTime)

    End Function

End Module


Public Class orderHistory

    Implements System.IComparable

    Private _tradeDate As DateTime
    Private _order As String
    Private _pair As String
    Private _price As Single
    Private _fee As Single
    Private _bonus As Single
    Private _size As Single

    '日付
    Public Property TradeDate() As DateTime
        Get
            Return Me._tradeDate
        End Get
        Set(ByVal value As DateTime)
            Me._tradeDate = value
        End Set
    End Property

    'Order
    Public Property Order() As String
        Get
            Return Me._order
        End Get
        Set(ByVal value As String)
            Me._order = value
        End Set
    End Property

    'Pair
    Public Property Pair() As String
        Get
            Return Me._pair
        End Get
        Set(ByVal value As String)
            Me._pair = value
        End Set

    End Property
    'BTC price
    Public Property Price() As Single
        Get
            Return Me._price
        End Get
        Set(ByVal value As Single)
            Me._price = value
        End Set
    End Property

    'fee
    Public Property Fee() As Single
        Get
            Return Me._fee
        End Get
        Set(ByVal value As Single)
            Me._fee = value
        End Set
    End Property

    'bonus
    Public Property Bonus() As Single
        Get
            Return Me._bonus
        End Get
        Set(ByVal value As Single)
            Me._bonus = value
        End Set
    End Property

    'BTC size
    Public Property Size() As Single
        Get
            Return Me._size
        End Get
        Set(ByVal value As Single)
            Me._size = value
        End Set
    End Property

    '日付チェック
    Public Function CompareTo(ByVal other As Object) As Integer Implements System.IComparable.CompareTo

        If Not (Me.GetType() = other.GetType()) Then
            Throw New ArgumentException()
        End If

        Dim obj As orderHistory = CType(other, orderHistory)
        Return Me._tradeDate.CompareTo(obj._tradeDate)

    End Function



End Class

