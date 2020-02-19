Imports System.Net.Sockets

Public Class PLC
    Private Client As TcpClient
    Private Stream As NetworkStream
    Private NodeAddress As String
    Public Sub New(ByVal IP As String, ByVal Port As Integer)
        Client = New TcpClient(IP, Port)
        Stream = Client.GetStream
        Stream.ReadTimeout = 1000
        Stream.WriteTimeout = 1000
        Dim Head As String = "46494E530000000C000000000000000000000000"
        Dim Length As Integer = Head.Length / 2 - 1
        Dim SendData(Length) As Byte
        Dim j As Integer = 1
        For i = 0 To Length
            SendData(i) = "&h" & Mid(Head, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Dim ReceiveData(24) As Byte
        Stream.Read(ReceiveData, 0, ReceiveData.Length)
        If ReceiveData(11) = 1 Then
            NodeAddress = Microsoft.VisualBasic.Right("0" & Hex(CLng(ReceiveData(19))), 2)
        End If
    End Sub

    Public Function readDM(ByVal startDM As Integer, ByVal number As Integer)
        Dim Head As String = "46494E530000001A0000000200000000"
        Dim ICF As String = "80"
        Dim RSV As String = "00"
        Dim GCT As String = "02"
        Dim DNA As String = "00"
        Dim DA1 As String = "00"
        Dim DA2 As String = "00"
        Dim SNA As String = "00"
        Dim SA1 As String = "00"
        Dim SA2 As String = "00"
        Dim SID As String = "00"
        Dim Length As String = Hex(number).ToString.PadLeft(4, "0")
        Dim Text = Hex(Val(startDM))
        Select Case Len(Text)
            Case 0 : Text = "0000"
            Case 1 : Text = "000" & Text
            Case 2 : Text = "00" & Text
            Case 3 : Text = "0" & Text
            Case Else : Text = Text
        End Select
        Dim Commend As String = Head & ICF & RSV & GCT & DNA & DA1 & DA2 & SNA & NodeAddress & SA2 & SID & "010182" & Text & "00" & Length
        Dim SendData(Commend.Length / 2 - 1) As Byte
        Dim j As Integer = 1
        For i = 0 To Commend.Length / 2 - 1
            SendData(i) = "&h" & Mid(Commend, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Dim ReceiveData(29 + 2 * number) As Byte
        Stream.Read(ReceiveData, 0, ReceiveData.Length)
        Dim Receive As String = Nothing
        For i = 0 To number * 2 - 1
            Receive = Receive & Hex(ReceiveData(30 + i)).ToString.PadLeft(2, "0")
        Next
        Return Receive
    End Function

    Public Function readBit(ByVal startDM As Integer, ByVal bit As Integer, ByVal number As Integer)
        Dim Head As String = "46494E530000001A0000000200000000"
        Dim ICF As String = "80"
        Dim RSV As String = "00"
        Dim GCT As String = "02"
        Dim DNA As String = "00"
        Dim DA1 As String = "00"
        Dim DA2 As String = "00"
        Dim SNA As String = "00"
        Dim SA1 As String = "00"
        Dim SA2 As String = "00"
        Dim SID As String = "00"
        'Dim Length As String = Hex(Amount).ToString.PadLeft(6, "0")
        Dim Text = Hex(Val(startDM))
        Select Case Len(Text)
            Case 0 : Text = "0000"
            Case 1 : Text = "000" & Text
            Case 2 : Text = "00" & Text
            Case 3 : Text = "0" & Text
            Case Else : Text = Text
        End Select
        Dim Text2 = Hex(Val(bit))
        Select Case Len(Text2)
            Case 0 : Text2 = "00"
            Case 1 : Text2 = "0" & Text2
            Case Else : Text2 = Text2
        End Select
        Dim Commend As String = Head & ICF & RSV & GCT & DNA & DA1 & DA2 & SNA & NodeAddress & SA2 & SID & "010102" & Text & Text2 & "0001"
        Dim SendData(Commend.Length / 2 - 1) As Byte
        Dim j As Integer = 1
        For i = 0 To Commend.Length / 2 - 1
            SendData(i) = "&h" & Mid(Commend, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Dim ReceiveData(29 + number) As Byte
        Stream.Read(ReceiveData, 0, ReceiveData.Length)
        Dim Receive As String = Nothing
        For i = 0 To number - 1
            Receive = Receive & Hex(ReceiveData(30 + i)).ToString
        Next
        Return Receive
    End Function
    Public Sub writeDM(ByVal DM As Integer, ByVal Data As String)
        Dim Head As String = "46494E530000001C0000000200000000"
        Dim ICF As String = "80"
        Dim RSV As String = "00"
        Dim GCT As String = "02"
        Dim DNA As String = "00"
        Dim DA1 As String = "00"
        Dim DA2 As String = "00"
        Dim SNA As String = "00"
        Dim SA1 As String = "00"
        Dim SA2 As String = "00"
        Dim SID As String = "00"
        Dim Text = Hex(Val(DM))
        Select Case Len(Text)
            Case 0 : Text = "0000"
            Case 1 : Text = "000" & Text
            Case 2 : Text = "00" & Text
            Case 3 : Text = "0" & Text
            Case Else : Text = Text
        End Select
        Dim Commend As String = Head & ICF & RSV & GCT & DNA & DA1 & DA2 & SNA & NodeAddress & SA2 & SID & "0102" & "82" & Text & Format(1, "000000") & Data.PadLeft(4, "0")
        Dim SendData(Commend.Length / 2 - 1) As Byte
        Dim j As Integer = 1
        For i = 0 To Commend.Length / 2 - 1
            SendData(i) = "&h" & Mid(Commend, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Threading.Thread.Sleep(100)
        Dim a As Integer = Client.Available
        Dim ReceiveData(a) As Byte
        Stream.Read(ReceiveData, 0, ReceiveData.Length)
    End Sub

    Public Sub writeBit(ByVal DM As Integer, ByVal Num As Integer, ByVal Data As String)
        Dim Head As String = "46494E530000001C0000000200000000"
        Dim ICF As String = "80"
        Dim RSV As String = "00"
        Dim GCT As String = "02"
        Dim DNA As String = "00"
        Dim DA1 As String = "00"
        Dim DA2 As String = "00"
        Dim SNA As String = "00"
        Dim SA1 As String = "00"
        Dim SA2 As String = "00"
        Dim SID As String = "00"
        Dim Text = Hex(Val(DM))
        Select Case Len(Text)
            Case 0 : Text = "0000"
            Case 1 : Text = "000" & Text
            Case 2 : Text = "00" & Text
            Case 3 : Text = "0" & Text
            Case Else : Text = Text
        End Select
        Dim Commend As String = Head & ICF & RSV & GCT & DNA & DA1 & DA2 & SNA & NodeAddress & SA2 & SID & "0102" & "31" & "00D4" & "00" & "0002" & "0101"
        Dim SendData(Commend.Length / 2 - 1) As Byte
        Dim j As Integer = 1
        For i = 0 To Commend.Length / 2 - 1
            SendData(i) = "&h" & Mid(Commend, j, 2)
            j = j + 2
        Next
        Stream.Write(SendData, 0, SendData.Length)
        Threading.Thread.Sleep(500)
        Dim ReceiveData(100) As Byte
        Dim a As Integer = Client.Available
        Stream.Read(ReceiveData, 0, 100)
    End Sub
    Sub Dispose()
        Stream.Dispose()
    End Sub
End Class