VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSmartSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *************************************************************************************************
' Copyright (C) Chris Waddell
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2, or (at your option)
' any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'
' Please consult the LICENSE.txt file included with this project for
' more details
'
' *************************************************************************************************
Option Explicit


' *************************************************************************************************
' A single entry in the send buffer. This stores information such as the data to send, the flags to
' use when sending it, whether to use Send or SendTo, if using SendTo then the remote host and port
' *************************************************************************************************
Private Type SendBufferEntry
    SendFlags       As MessageCommunicationFlags    ' The flags used when sending the data
    SendData()      As Byte                         ' The actual data
    RemoteHost      As CIP4Address                  ' The remote host (only used with SendTo)
    RemotePort      As Long                         ' The remote port (only used with SendTo)
    UseSendTo       As Boolean                      ' Whether to use SendTo or Send
    FlagDeletion    As Boolean                      ' This flag is set when the buffer is empty
End Type


' *************************************************************************************************
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Exception As CWinsockException)

' *************************************************************************************************
' Raised whenever data is partially sent.
' *************************************************************************************************
Public Event OnSendProgress(BytesSent As Long, BytesRemaining As Long)

' *************************************************************************************************
' Raised whenever all data is sent.
' *************************************************************************************************
Public Event OnSendComplete()


' *************************************************************************************************
' Occurs whenever new data arrives
' ********************************.*****************************************************************
Public Event OnDataArrive(BytesRecieved As Long, BytesTotal As Long)

' The internal send buffer
Private m_SendBuffer() As SendBufferEntry


' The internal socket class.
Private WithEvents m_Socket As CSocket
Attribute m_Socket.VB_VarHelpID = -1


' *************************************************************************************************
' This must be set before attempting to use the class to send data.
' *************************************************************************************************
Public Property Set SocketObject(Value As CSocket)
    Set m_Socket = Value
End Property


' *************************************************************************************************
' Returns the total bytes waiting to be sent on the socket.
' *************************************************************************************************
Public Property Get SendBufferSize() As Long

  Dim i As Integer
  Dim BufferSize As Long
  
    ' Loop through each send buffer entry and add up the size of the data contained in it
    For i = 1 To UBound(m_SendBuffer)
        BufferSize = BufferSize + UBound(m_SendBuffer(i).SendData)
    Next i
    
    SendBufferSize = BufferSize

End Property


' *************************************************************************************************
' Description: Adds a new entry to the send buffer list for sending.
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    13/04/2004 by Chris Waddell
' *************************************************************************************************
Private Function AddNewSendBuffer(Buffer() As Byte, Flags As MessageCommunicationFlags, RemoteHost As CIP4Address, RemotePort As Long, UseSendTo As Boolean)

    ' Add a new entry to the send buffer
    ReDim Preserve m_SendBuffer(UBound(m_SendBuffer) + 1)
    
    ' Save the remote host
    Set m_SendBuffer(UBound(m_SendBuffer)).RemoteHost = RemoteHost
    
    ' Save the remote port
    m_SendBuffer(UBound(m_SendBuffer)).RemotePort = RemotePort
    
    ' Save whether to use SendTo or Send
    m_SendBuffer(UBound(m_SendBuffer)).UseSendTo = UseSendTo
    
    ' Resize the new send buffer
    ReDim m_SendBuffer(UBound(m_SendBuffer)).SendData(UBound(Buffer) + 1)
    
    ' Copy the data to send into it
    RtlMoveMemory ByVal VarPtr(m_SendBuffer(UBound(m_SendBuffer)).SendData(1)), ByVal VarPtr(Buffer(0)), ByVal UBound(Buffer) + 1
    
End Function


' *************************************************************************************************
' Description: Attempts to send all data in the send buffer. It will return false if an error
'              occured or true otherwise. If the data cannot be sent at this moment in time
'              (i.e. WSAEWOULDBLOCK occurs, or the number of bytes sent is less than the desired
'              amount then it still returns true, but waits until later when an FD_WRITE is issued
'              before trying again. Note that an error could occur at that stage, so this method
'              returning true does not necessarily mean that the data is or will be sent
'              successfuly.
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    13/04/2004 by Chris Waddell
' *************************************************************************************************
Private Function SendBufferedData() As Boolean

  Dim i                           As Integer
  Dim SockAddr            As API_SOCKADDR_IN
  Dim lngRetVal           As Long
  Dim BufferCleared   As Boolean

    ' If there is no data to send, just return true and do nothing
    If UBound(m_SendBuffer) = 0 Then
        SendBufferedData = True
        Exit Function
    End If
    
    ' Otherwise loop through each of the send buffers
    For i = 1 To UBound(m_SendBuffer)

        ' Check whether to use Send or SendTo
        If m_SendBuffer(i).UseSendTo Then
        
            ' When using SendTo we need to fill out the socket address
            SockAddr.sin_addr.S_addr = m_SendBuffer(i).RemoteHost.Address
            SockAddr.sin_family = m_Socket.AddressFamily
            SockAddr.sin_port = api_htons(UnsignedToInteger(m_SendBuffer(i).RemotePort))
            
            ' Try to send the data
            lngRetVal = api_SendTo(m_Socket.SocketHandle, m_SendBuffer(i).SendData(1), UBound(m_SendBuffer(i).SendData), m_SendBuffer(i).SendFlags, SockAddr, LenB(SockAddr))
        Else
            ' Try to send the data
            lngRetVal = api_Send(m_Socket.SocketHandle, m_SendBuffer(i).SendData(1), UBound(m_SendBuffer(i).SendData), m_SendBuffer(i).SendFlags)
        End If
        
        ' If an error occured that was not WSAEWOULDBLOCK then return false, otherwise just exit.
        If lngRetVal = SOCKET_ERROR Then
            SendBufferedData = Err.LastDllError = WSAEWOULDBLOCK
            Exit Function
        
        ' If no error occured, but the number of bytes sent was less than the desired amount
        ' then remove the bytes that were sent from the buffer, then and stop trying to send any
        ' more data.
        ElseIf lngRetVal < UBound(m_SendBuffer(i).SendData) Then
            RtlMoveMemory ByVal VarPtr(m_SendBuffer(i).SendData(1)), ByVal VarPtr(m_SendBuffer(i).SendData(lngRetVal + 1)), ByVal UBound(m_SendBuffer(i).SendData) - lngRetVal
            ReDim Preserve m_SendBuffer(i).SendData(UBound(m_SendBuffer(i).SendData) - lngRetVal)
            RaiseEvent OnSendProgress(lngRetVal, SendBufferSize)
            Exit For
        
        ' If the data was all successfuly sent then flag this buffer fot deletion
        Else
            m_SendBuffer(i).FlagDeletion = True
        End If
    Next i
    
    Do While Not BufferCleared
        BufferCleared = True
    
        ' Now loop through each buffer
        For i = 1 To UBound(m_SendBuffer)
        
            ' If it has been flagged for deletion then
            If m_SendBuffer(i).FlagDeletion Then
            
                ' Swap it with the last one and resize the send buffer down by 1 entry
                m_SendBuffer(i) = m_SendBuffer(UBound(m_SendBuffer))
                ReDim Preserve m_SendBuffer(UBound(m_SendBuffer) - 1)

                                ' Check the buffers again
                BufferCleared = False
                Exit For
            End If
        Next i
    Loop
    
    ' If the send buffer size is now 0 then we've sent all the data so raise the event
    If UBound(m_SendBuffer) = 0 Then
        RaiseEvent OnSendComplete
    End If

    ' If we get to here it was successful
    SendBufferedData = True

End Function



' *************************************************************************************************
' Description: Attempts to send the data. It will return false if an error occured or true
'              otherwise. If the data cannot be sent at this moment in time (i.e. WSAEWOULDBLOCK
'              occurs, or the number of bytes sent is less than the desired amount then it still
'              returns true, but waits until later when an FD_WRITE is issued
'              before trying again. Note that an error could occur at that stage, so this method
'              returning true does not necessarily mean that the data is or will be sent
'              successfuly.
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    13/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Send(Data As Variant, Optional Flags As MessageCommunicationFlags = Msg_None, Optional MaxSize As Long) As Boolean

  Dim ErrorObject   As CWinsockException
  Dim NumBuffers    As Long
  Dim Buffer()      As Byte
  Dim TempBuffer()  As Byte
  Dim BytesLeft     As Long
  Dim BufferSize    As Long
  Dim i             As Integer

    ' Check that the socket object has actually been set
    If m_Socket Is Nothing Then
        Err.Raise -1, "CSmartCommunicator.Send", "The socket object is nothing"
        Exit Function
    End If
    
    Buffer = GetBufferByVariant(Data)
    
    ' Calculate the size of the buffer to send
    BufferSize = (UBound(Buffer) - LBound(Buffer) + 1)
    
    ' Calculate the number of send buffers required
    If MaxSize <= 0 Then MaxSize = BufferSize
    NumBuffers = BufferSize / MaxSize
    
    ' Add each send buffer
    For i = 0 To NumBuffers - 1
        ReDim TempBuffer(MaxSize - 1)
        RtlMoveMemory ByVal VarPtr(TempBuffer(0)), ByVal VarPtr(Buffer(i * MaxSize)), MaxSize
        AddNewSendBuffer TempBuffer, Flags, Nothing, 0, False
    Next i
    
    ' Calculate if there are any bytes left which don't fill the max size
    BytesLeft = BufferSize - (MaxSize * NumBuffers)
    
    ' If so then add those too
    If BytesLeft > 0 Then
        ReDim TempBuffer(BytesLeft - 1)
        RtlMoveMemory ByVal VarPtr(TempBuffer(0)), ByVal VarPtr(Buffer(BufferSize - BytesLeft)), BytesLeft
        AddNewSendBuffer TempBuffer, Flags, Nothing, 0, False
    End If

    ' Try to send the buffered data. If an error occured then raise it otherwise return true.
    If Not SendBufferedData Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSmartCommunicator.Send"
        RaiseEvent OnError(ErrorObject)
    Else
        Send = True
    End If

End Function


' *************************************************************************************************
' Description: Attempts to send the data. It will return false if an error occured or true
'              otherwise. If the data cannot be sent at this moment in time (i.e. WSAEWOULDBLOCK
'              occurs, or the number of bytes sent is less than the desired amount then it still
'              returns true, but waits until later when an FD_WRITE is issued
'              before trying again. Note that an error could occur at that stage, so this method
'              returning true does not necessarily mean that the data is or will be sent
'              successfuly.
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    13/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function SendTo(Data As Variant, RemoteHost As CIP4Address, RemotePort As Long, Optional Flags As MessageCommunicationFlags = Msg_None, Optional MaxSize As Long) As Boolean
  
  Dim ErrorObject   As CWinsockException
  Dim NumBuffers    As Long
  Dim Buffer()      As Byte
  Dim TempBuffer()  As Byte
  Dim BytesLeft     As Long
  Dim BufferSize    As Long
  Dim i             As Integer

    ' Check that the socket object has actually been set
    If m_Socket Is Nothing Then
        Err.Raise -1, "CSmartCommunicator.SendTo", "The socket object is nothing"
        Exit Function
    End If
    
    ' Make sure there is a remote host
    If RemoteHost Is Nothing Then
        Err.Raise -1, "CSmartCommunicator.SendTo", "No remote host parameter was passed"
        Exit Function
    End If
    
    ' Make sure there is a remote port
    If RemotePort <= 0 Then
        Err.Raise -1, "CSmartCommunicator.SendTo", "No remote port parameter was passed"
        Exit Function
    End If

    Buffer = GetBufferByVariant(Data)

    ' Calculate the size of the buffer to send
    BufferSize = (UBound(Buffer) - LBound(Buffer) + 1)
    
    ' Calculate the number of send buffers required
    If MaxSize <= 0 Then MaxSize = BufferSize
    NumBuffers = BufferSize / MaxSize
    
    ' Add each send buffer
    For i = 0 To NumBuffers - 1
        ReDim TempBuffer(MaxSize - 1)
        RtlMoveMemory ByVal VarPtr(TempBuffer(0)), ByVal VarPtr(Buffer(i * MaxSize)), MaxSize
        AddNewSendBuffer TempBuffer, Flags, RemoteHost, RemotePort, True
    Next i
    
    ' Calculate if there are any bytes left which don't fill the max size
    BytesLeft = BufferSize - (MaxSize * NumBuffers)
    
    ' If so then add those too
    If BytesLeft > 0 Then
        ReDim TempBuffer(BytesLeft - 1)
        RtlMoveMemory ByVal VarPtr(TempBuffer(0)), ByVal VarPtr(Buffer(BufferSize - BytesLeft)), BytesLeft
        AddNewSendBuffer TempBuffer, Flags, RemoteHost, RemotePort, True
    End If

    ' Try to send the buffered data. If an error occured then raise it otherwise return true.
    If Not SendBufferedData Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSmartCommunicator.Send"
        RaiseEvent OnError(ErrorObject)
    Else
        SendTo = True
    End If

End Function


' *************************************************************************************************
' This occurs when the socket is ready for writing again. That is, whenever a previous call to
' Send or SendTo didn't complete or was unsuccessful due to lack of internal buffer space.
' This event indicates that there is now buffer space available and we can try sending again now.
' *************************************************************************************************
Private Sub m_Socket_ReadyForWriting()

  Dim ErrorObject As CWinsockException

    ' So try to send the buffered data
    If Not SendBufferedData Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSmartCommunicator.m_Socket_ReadyForWriting"
        RaiseEvent OnError(ErrorObject)
    End If
    
End Sub


Private Sub Class_Initialize()
    ReDim m_SendBuffer(0)
    ReDim m_RecieveBuffer(0)
End Sub


Private Sub Class_Terminate()
    Erase m_SendBuffer
    Set m_Socket = Nothing
End Sub


