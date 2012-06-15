''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command2112 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command2112.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Public Class Network_MT

	Private Shared objQSync As New Object
	Private Shared RequestQueue As New Generic.Queue(Of ReadWriteKeyValuePair(Of String, Integer))
	Private Shared booStop, booOkayToSend As Boolean
	Private Shared mre As New ManualResetEvent(False)
	Private Shared mre2 As New ManualResetEvent(True)
	Private Shared s As TcpClient
	Private Shared t As Thread
	Private Shared lBuffer As New Generic.List(Of Byte)
	Public Shared LogRawBytes As Boolean
	Private Shared WithEvents tmr As New System.Timers.Timer

	Public Shared Sub Start()

		t = New Thread(AddressOf Process)
		t.Name = "Denon Message Processing Thread - Network_MT.Process()"
		s = Nothing
		mre.Reset()
		booOkayToSend = True

		If t.ThreadState = ThreadState.Unstarted OrElse t.ThreadState = ThreadState.Stopped Then
			t.Start()
		End If

		AddHandler Commands.RequestReady, AddressOf EnqueueRequest

	End Sub

	Public Shared Sub [Stop]()

		RemoveHandler Commands.RequestReady, AddressOf EnqueueRequest

		If Not IsNothing(t) AndAlso (t.ThreadState = Threading.ThreadState.Running OrElse t.ThreadState = ThreadState.WaitSleepJoin) Then
			tmr.Stop()
			booOkayToSend = False
			booStop = True

			If Not mre.WaitOne(5000, False) Then
				PostMessage("WaitOne expired after 5 seconds while waiting for Network_MT.Process() to stop its work and signal completion.", Nothing)
			End If
		End If

	End Sub

	Public Shared Sub ResetConnection()
		[Stop]()
		Start()
	End Sub

	Private Shared Sub Process()

		Dim bs(), br(255) As Byte
		Dim l As Generic.List(Of String)
		Dim strSendCommand As String
		Dim mq As New Queue(Of String)
		Dim kvp As ReadWriteKeyValuePair(Of String, Integer)
		Dim intTO, intRead As Integer

		Try

			Connect()

			Do

				Try

					If booOkayToSend Then

						SyncLock objQSync

							If RequestQueue.Count > 0 Then
								kvp = RequestQueue.Dequeue
							Else
								kvp = Nothing
							End If

						End SyncLock

						If Not IsNothing(kvp) AndAlso Not String.IsNullOrEmpty(kvp.Key) Then

							strSendCommand = kvp.Key
							intTO = kvp.Value

							PostMessage("SEND: " & strSendCommand, Nothing)

							bs = System.Text.Encoding.UTF8.GetBytes(strSendCommand & vbCr)

							Try
								s.GetStream.Write(bs, 0, bs.Length)
							Catch ex As Exception
								If ex.Message.Contains("period of time") Then
									PostMessage("TIMEOUT ON SEND: " & strSendCommand, Nothing)
								Else
									PostMessage("EXCEPTION ON SEND: " & strSendCommand, New Exception("An exception occured while sending a message to the Denon", ex))
								End If
							End Try

							booOkayToSend = False
							tmr.Interval = intTO
							tmr.Start()

						End If

					End If

					l = New Generic.List(Of String)

					Try

						If s.Available > 0 Then

							intRead = s.GetStream.Read(br, 0, br.Length)

							If intRead > 0 Then
								l.AddRange(ParseCommandResponse(intRead, br))
								Array.Clear(br, 0, intRead)
							End If

						End If

						Threading.Thread.Sleep(1)

					Catch ex As Exception
						If ex.Message.Contains("period of time") Then
							PostMessage("TIMEOUT ON RECV", Nothing)
						Else
							PostMessage("EXCEPTION ON RECV", New Exception("An exception occured while receiving a message from the Denon", ex))
						End If
					End Try

					For Each strX As String In l
						PostMessage("RECV: " & strX, Nothing)
					Next

					PostResponse(l)

					If booStop Then
						booStop = False
						Exit Do
					End If

				Catch ex As Exception
					PostMessage("EXCEPTION", New Exception("An exception occured in the Denon messaging loop.", ex))
				End Try

			Loop

			Disconnect()

		Catch ex As Exception
			PostMessage("EXCEPTION", New Exception("An exception occured in the Denon process routine.", ex))
		Finally
			mre.Set()
		End Try

	End Sub

	Private Shared Sub PostMessage(ByVal m As String, ByVal ex As Exception)
		RaiseEvent MessageReady(New ReadWriteKeyValuePair(Of String, Exception)(Helper.GetNowString & " - " & m, ex))
	End Sub

	Private Shared Sub PostResponse(ByVal l As Generic.List(Of String))
		RaiseEvent ResponseReady(l)
	End Sub

	Private Shared Sub EnqueueRequest(ByVal kvp As ReadWriteKeyValuePair(Of String, Integer))
		SyncLock objQSync
			RequestQueue.Enqueue(kvp)
		End SyncLock
	End Sub

	Public Shared Event ResponseReady(ByVal l As Generic.List(Of String))
	Public Shared Event MessageReady(ByVal kvp As ReadWriteKeyValuePair(Of String, Exception))

	Private Shared Function ParseCommandResponse(ByVal BytesRead As Integer, ByVal byt As Byte()) As Generic.List(Of String)

		Dim lByt As Generic.List(Of Byte) = New Generic.List(Of Byte)(byt).GetRange(0, BytesRead)
		Dim lCmdB As New Generic.List(Of Byte)
		Dim lCmdS As New Generic.List(Of String)
		Dim strX As String
		Dim b As Byte

		If LogRawBytes Then
			DebugToFile(lByt.ToArray)
		End If

		If lBuffer.Count <> 0 Then
			lCmdB.AddRange(lBuffer)
			lBuffer.Clear()
		End If

		For intX As Integer = 0 To lByt.Count - 1

			b = lByt(intX)

			If b <> 13 Then
				lCmdB.Add(b)
			Else
				strX = System.Text.Encoding.UTF8.GetString(lCmdB.ToArray)
				lCmdB.Clear()

				If Not String.IsNullOrEmpty(strX) Then
					lCmdS.Add(strX)
				End If

			End If

		Next

		If lCmdB.Count <> 0 Then
			lBuffer.AddRange(lCmdB)
		End If

		Return lCmdS

	End Function

	Private Shared Sub DebugToFile(ByVal byt As Byte())

		Dim fs As New IO.FileStream(My.Application.Info.DirectoryPath & "\debug.bin", IO.FileMode.Append)
		Dim bw As New IO.BinaryWriter(fs)

		bw.Write(byt)
		bw.Close()

	End Sub

	Private Shared Sub Connect()

		Dim awh As WaitHandle
		Dim ar As IAsyncResult

		s = New TcpClient
		s.ReceiveTimeout = 5000
		s.SendTimeout = 5000

		PostMessage("Attempting to open connection with " & My.CustomSettings.IPAddress.Instance.Value, Nothing)

		ar = s.BeginConnect(IPAddress.Parse(My.CustomSettings.IPAddress.Instance.Value), 23, Nothing, Nothing)

		awh = ar.AsyncWaitHandle

		If Not awh.WaitOne(5000, False) Then
			s.Close()
			PostMessage("Connection could not be established with " & My.CustomSettings.IPAddress.Instance.Value, Nothing)
			Throw New TimeoutException()
		End If

		s.EndConnect(ar)

		PostMessage("Connection established with " & My.CustomSettings.IPAddress.Instance.Value, Nothing)

	End Sub

	Private Shared Sub Disconnect()

		Dim strIP As String

		If Not IsNothing(s) Then

			strIP = DirectCast(s.Client.RemoteEndPoint, IPEndPoint).Address.ToString

			PostMessage("Attempting to close connection with " & strIP, Nothing)

			s.Client.Shutdown(SocketShutdown.Both)
			s.Client.Close()
			s.Close()

			PostMessage("Connection closed with " & strIP, Nothing)

			s = Nothing

		End If

	End Sub

	Private Shared Sub tmr_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles tmr.Elapsed
		tmr.Stop()
		booOkayToSend = True
	End Sub
End Class
