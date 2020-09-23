VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocketCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' *************************************************************************************************
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Socket As CSocket, Exception As CWinsockException)


' *************************************************************************************************
' Socket ready for accepting a new incoming connection.
' *************************************************************************************************
Public Event ConnectionRequest(Socket As CSocket)


' *************************************************************************************************
' Connection identified by socket has been closed.
' *************************************************************************************************
Public Event Closed(Socket As CSocket)


' *************************************************************************************************
' Connection or multipoint join operation initiated on socket completed.
' *************************************************************************************************
Public Event Connected(Socket As CSocket)


' *************************************************************************************************
' Socket ready for reading.
' *************************************************************************************************
Public Event DataArrived(Socket As CSocket)


' *************************************************************************************************
' Socket ready for writing.
' *************************************************************************************************
Public Event ReadyForWriting(Socket As CSocket) 'BytesSent As Long)


' *************************************************************************************************
' Raised whenever data is partially sent.
' *************************************************************************************************
Public Event OnSendProgress(Socket As CSocket, BytesSent As Long, BytesRemaining As Long)


' *************************************************************************************************
' Raised whenever all data is sent.
' *************************************************************************************************
Public Event OnSendComplete(Socket As CSocket)


Private WithEvents EventWindow As CWindow
Attribute EventWindow.VB_VarHelpID = -1

Private m_Sockets As Collection
Private m_Senders As Collection


Public Sub Add(Socket As CSocket)
    
  Dim SmartSender As CSmartSender
    
    If Socket Is Nothing Then Exit Sub
    
    ' Disable the current event handler
    Socket.SetBlocking False
    
    ' Add our own event handler
    Socket.SetBlocking True, EventWindow
    
    ' Add the socket to the collection
    m_Sockets.Add Socket, Socket.SocketHandle
    
    ' Create a new smart sender
    Set SmartSender = New CSmartSender
    
    ' Add the socket object to it
    Set SmartSender.SocketObject = Socket
    
    ' Add it to the collection
    m_Senders.Add SmartSender, Socket.SocketHandle
    
End Sub


Public Sub Remove(Key As Variant)
    m_Sockets.Remove Key
End Sub


Private Sub Class_Initialize()
    Set EventWindow = New CWindow
End Sub


Private Sub Class_Terminate()
    Set EventWindow = Nothing
End Sub


Private Sub EventWindow_MessageArrived(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, lngRetVal As Long, bHandled As Boolean)

  Dim lngErrorCode          As Long
  Dim lngEvent              As Long
  Dim lngBytesRecieved      As Long
  Dim ErrorObject           As CWinsockException
  Dim Buffer()              As Byte
  Dim Socket                As CSocket
  Dim i                     As Integer

    ' Extract the error code and event from the lParam
    lngErrorCode = HiWord(lParam)
    lngEvent = LoWord(lParam)
    
    
    ' Get the socket object
    For i = 1 To m_Sockets.Count
        If m_Sockets(i).SocketHandle = wParam Then
            Set Socket = m_Sockets(i)
            Exit For
        End If
    Next i
    
    ' Check for errors
    If lngErrorCode <> 0 Then
        Set ErrorObject = New CWinsockException
        ErrorObject.ErrorCode = lngErrorCode
        ErrorObject.Source = "CSocket.EventWindow_MessageArrived"
        RaiseEvent OnError(Socket, ErrorObject)
        Exit Sub
    End If

    ' Check what the event is
    Select Case lngEvent
        Case AsyncSelectFlagsType.ASFlag_Accept
            RaiseEvent ConnectionRequest(Socket)
        Case AsyncSelectFlagsType.ASFlag_Close
            Socket.CloseSocket
            RaiseEvent Closed(Socket)
        Case AsyncSelectFlagsType.ASFlag_Connect
            RaiseEvent Connected(Socket)
        Case AsyncSelectFlagsType.ASFlag_Read
            RaiseEvent DataArrived(Socket)
        Case AsyncSelectFlagsType.ASFlag_Write
            RaiseEvent ReadyForWriting(Socket)
        Case Else
            Debug.Print "Unknown event"
    End Select

    ' Pass the message on to the socket so that the smart sender will actually see the
    ' events occuring
    Socket.EventWindow_MessageArrived hwnd, uMsg, wParam, lParam, lngRetVal, bHandled

End Sub
