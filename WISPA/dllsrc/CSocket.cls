VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
' Protocol types which use IP
' *************************************************************************************************
Public Enum IPProtocolType
    Proto_IP = IPPROTO_IP                    ' dummy for IP
    Proto_Icmp = IPPROTO_ICMP                ' control message protocol
    Proto_Igmp = IPPROTO_IGMP                ' group management protocol
    Proto_Ggp = IPPROTO_GGP                  ' gateway^2 (deprecated)
    Proto_Tcp = IPPROTO_TCP                  ' tcp
    Proto_Pup = IPPROTO_PUP                  ' pup
    Proto_Udp = IPPROTO_UDP                  ' user datagram protocol
    Proto_Idp = IPPROTO_IDP                  ' xns idp
    Proto_ND = IPPROTO_ND                    ' UNOFFICIAL net disk proto
    Proto_Raw = IPPROTO_RAW                  ' raw IP packet
    Proto_Unspecified = 0                    ' unspecified
    Proto_Ipx = IPPROTO_IPX                  ' IPX protocol
    Proto_Spx = IPPROTO_SPX                  ' SPX protocol
    Proto_SpxII = IPPROTO_SPXII              ' SPXII protocol
    Proto_Unknown = -1                       ' unknown protocol type
    Proto_Max = IPPROTO_MAX
End Enum


' Socket shutdown types
' *************************************************************************************************
Public Enum SocketShutdownType

    ' If the how parameter is SD_RECEIVE, subsequent calls to the recv function on the socket
    ' will be disallowed. This has no effect on the lower protocol layers. For TCP sockets,
    ' if there is still data queued on the socket waiting to be received, or data arrives
    ' subsequently, the connection is reset, since the data cannot be delivered to the user. For
    ' UDP sockets, incoming datagrams are accepted and queued. In no case will an ICMP error
    ' packet be generated.
    SDType_Recieve = SD_RECEIVE
    
    ' If the how parameter is SD_SEND, subsequent calls to the send function are disallowed. For
    ' TCP sockets, a FIN will be sent after all data is sent and acknowledged by the receiver.
    SDType_Send = SD_SEND
    
    ' Setting how to SD_BOTH disables both sends and receives as described above.
    SDType_Both = SD_BOTH
End Enum


' *************************************************************************************************
' Flags for sending and recieving data
' *************************************************************************************************
Public Enum MessageCommunicationFlags
    Msg_None = 0
    Msg_OutOfBand = MSG_OOB
    Msg_PeekData = MSG_PEEK
    Msg_NoRoutingTables = MSG_DONTROUTE
    Msg_PartialSendRecv = MSG_PARTIAL
End Enum


' *************************************************************************************************
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Exception As CWinsockException)


' *************************************************************************************************
' Socket ready for accepting a new incoming connection.
' *************************************************************************************************
Public Event ConnectionRequest()


' *************************************************************************************************
' Connection identified by socket has been closed.
' *************************************************************************************************
Public Event Closed()


' *************************************************************************************************
' Connection or multipoint join operation initiated on socket completed.
' *************************************************************************************************
Public Event Connected()


' *************************************************************************************************
' Socket ready for reading.
' *************************************************************************************************
Public Event DataArrived()


' *************************************************************************************************
' Socket ready for writing.
' *************************************************************************************************
Public Event ReadyForWriting() 'BytesSent As Long)


' *************************************************************************************************
' Used when checking the socket for readabiloty, writability or errors using api_Select.
' *************************************************************************************************
Public Enum SocketCheckFlags
    BlockIndefinetly            ' Block indefinetly. Pass a null pointer as API_TIMEVAL.
    ReturnImmediately           ' Used for polling the socket. Pass a timeout of 0.
    UseTimeout                  ' Use a timeout time.
End Enum


' The socket handle.
Private m_lngSocketHandle           As Long

' Socket creation parameters
Private m_AddressFamily             As AddressFamilyType
Private m_SocketType                As SocketType
Private m_Protocol                  As IPProtocolType

Private m_Async                     As Boolean


' Socket input output controller and socket options
Private WithEvents m_SocketIO       As CSocketIO
Attribute m_SocketIO.VB_VarHelpID = -1
Private WithEvents m_SocketOptions  As CSocketOptions
Attribute m_SocketOptions.VB_VarHelpID = -1

' Event window
Private WithEvents EventWindow      As CWindow
Attribute EventWindow.VB_VarHelpID = -1



' *************************************************************************************************
' Socket blocking.
' *************************************************************************************************
Public Property Get IsBlocking() As Boolean
    IsBlocking = m_Async
End Property

Public Function SetBlocking(Value As Boolean, Optional ByVal Handler As CWindow = Nothing)

  Dim ErrorObject As CWinsockException

    If Handler Is Nothing Then Set Handler = EventWindow

    If Not Value And m_Async Then
        If api_WSAAsyncSelect(m_lngSocketHandle, Handler.WindowHandle, Handler.WinsockMessage, _
                           FD_READ Or FD_WRITE Or FD_OOB Or FD_ACCEPT Or FD_CONNECT Or _
                           FD_CLOSE) = SOCKET_ERROR Then
            Set ErrorObject = New CWinsockException
            ErrorObject.Source = "CSocket.Let Blocking"
            RaiseEvent OnError(ErrorObject)
            Exit Function
        End If
    ElseIf Not m_Async Then
        If api_WSAAsyncSelect(m_lngSocketHandle, Handler.WindowHandle, Handler.WinsockMessage, 0) = SOCKET_ERROR Then
            Set ErrorObject = New CWinsockException
            ErrorObject.Source = "CSocket.Let Blocking"
            RaiseEvent OnError(ErrorObject)
            Exit Function
        End If
    End If
    
    m_Async = Value
    
End Function


' *************************************************************************************************
' Socket handle.
' *************************************************************************************************
Public Property Get SocketHandle() As Long
    SocketHandle = m_lngSocketHandle
End Property

Public Property Let SocketHandle(Value As Long)
    ' Only let them change the socket handle if it's already not in use
    If m_lngSocketHandle = INVALID_SOCKET Then m_lngSocketHandle = Value
End Property


' *************************************************************************************************
' Address Family.
' *************************************************************************************************
Public Property Get AddressFamily() As AddressFamilyType
    AddressFamily = m_AddressFamily
End Property


' *************************************************************************************************
' Socket Type.
' *************************************************************************************************
Public Property Get SocketType() As SocketType
    SocketType = m_SocketType
End Property



' *************************************************************************************************
' Protocol.
' *************************************************************************************************
Public Property Get Protocol() As IPProtocolType
    Protocol = m_Protocol
End Property


' *************************************************************************************************
' Socket input output controller.
' *************************************************************************************************
Public Property Get InputOutputControl() As CSocketIO
    Set InputOutputControl = m_SocketIO
End Property


' *************************************************************************************************
' Socket options.
' *************************************************************************************************
Public Property Get Options() As CSocketOptions
    Set Options = m_SocketOptions
End Property


' *************************************************************************************************
' Properties to assist with human readable output.
' *************************************************************************************************

Public Property Get AddressFamilyName(af As AddressFamilyType)
    AddressFamilyName = GetAddressFamilyName(af)
End Property

Public Property Get SocketTypeName(st As SocketType)
    SocketTypeName = GetSocketTypeName(st)
End Property

Public Property Get ProtocolName(Protocol As IPProtocolType)
    ProtocolName = GetProtocolName(Protocol)
End Property


' *************************************************************************************************
' Returns true if the socket is writable, false otherwise.
' *************************************************************************************************
Public Property Get IsSocketWritable() As Boolean

  Dim WriteInfo     As API_FD_SET
  Dim TimeOut       As API_TIMEVAL
  Dim lngRetVal     As Long
  Dim ErrorObject   As CWinsockException

    TimeOut.tv_sec = 0
    TimeOut.tv_usec = 0
    
    WriteInfo.fd_array(0) = m_lngSocketHandle
    WriteInfo.fd_count = 1
    
    lngRetVal = api_Select(ByVal 0&, WriteInfo, ByVal 0&, ByVal 0&, ByVal VarPtr(TimeOut))

    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.IsSocketWritable"
        RaiseEvent OnError(ErrorObject)
    Else
        IsSocketWritable = CBool(lngRetVal)
    End If
    
End Property


' *************************************************************************************************
' Returns true if the socket is readable, false otherwise.
' *************************************************************************************************
Public Property Get IsSocketReadable() As Boolean

  Dim ReadInfo      As API_FD_SET
  Dim TimeOut       As API_TIMEVAL
  Dim lngRetVal     As Long
  Dim ErrorObject   As CWinsockException

    TimeOut.tv_sec = 0
    TimeOut.tv_usec = 0
    
    ReadInfo.fd_array(0) = m_lngSocketHandle
    ReadInfo.fd_count = 1
    
    lngRetVal = api_Select(ByVal 0&, ByVal 0&, ReadInfo, ByVal 0&, ByVal VarPtr(TimeOut))

    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.IsSocketReadable"
        RaiseEvent OnError(ErrorObject)
    Else
        IsSocketReadable = CBool(lngRetVal)
    End If
    
End Property


' *************************************************************************************************
' Returns true if the socket is in an erronous state, false otherwise.
' *************************************************************************************************
Public Property Get IsSocketErronous() As Boolean

  Dim ErrorInfo     As API_FD_SET
  Dim TimeOut       As API_TIMEVAL
  Dim lngRetVal     As Long
  Dim ErrorObject   As CWinsockException

    TimeOut.tv_sec = 0
    TimeOut.tv_usec = 0
    
    ErrorInfo.fd_array(0) = m_lngSocketHandle
    ErrorInfo.fd_count = 1
    
    lngRetVal = api_Select(ByVal 0&, ByVal 0&, ByVal 0&, ErrorInfo, ByVal VarPtr(TimeOut))

    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.IsSocketErronous"
        RaiseEvent OnError(ErrorObject)
    Else
        IsSocketErronous = CBool(lngRetVal)
    End If
    
End Property


' *************************************************************************************************
' Local host address
' *************************************************************************************************
Public Property Get LocalHost() As CDnsResult

  Dim SockAddr      As API_SOCKADDR_IN
  Dim ErrorObject   As CWinsockException
  Dim Dns           As CDns
    
    If m_lngSocketHandle = INVALID_SOCKET Then Exit Property
    
    If api_GetSockName(m_lngSocketHandle, SockAddr, LenB(SockAddr)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.LocalHost"
        RaiseEvent OnError(ErrorObject)
    Else
        Set Dns = New CDns
        Set LocalHost = Dns.Resolve(GetStrIPFromLong(SockAddr.sin_addr.S_addr))
        Set Dns = Nothing
    End If
    
    ' If an error occured then just return a new blank dns result
    If LocalHost Is Nothing Then
        Set LocalHost = New CDnsResult
    End If
    
End Property


' *************************************************************************************************
' Remote host address
' *************************************************************************************************
Public Property Get RemoteHost() As CDnsResult

  Dim SockAddr As API_SOCKADDR_IN
  Dim ErrorObject As CWinsockException
  Dim Dns As CDns
    
    If m_lngSocketHandle = INVALID_SOCKET Then Exit Property
    
    If api_GetPeerName(m_lngSocketHandle, SockAddr, LenB(SockAddr)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.RemoteHost"
        RaiseEvent OnError(ErrorObject)
    Else
        Set Dns = New CDns
        Set RemoteHost = Dns.Resolve(GetStrIPFromLong(SockAddr.sin_addr.S_addr))
        Set Dns = Nothing
    End If

    ' If an error occured then just return a new blank dns result
    If RemoteHost Is Nothing Then
        Set RemoteHost = New CDnsResult
    End If

End Property


' *************************************************************************************************
' Local port.
' *************************************************************************************************
Public Property Get LocalPort() As Long

  Dim SockAddr As API_SOCKADDR_IN
  Dim ErrorObject As CWinsockException

    If m_lngSocketHandle = INVALID_SOCKET Then Exit Property

    If api_GetSockName(m_lngSocketHandle, SockAddr, LenB(SockAddr)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.LocalHost"
        RaiseEvent OnError(ErrorObject)
    Else
        LocalPort = IntegerToUnsigned(api_ntohs(SockAddr.sin_port))
    End If

End Property


' *************************************************************************************************
' Remote port.
' *************************************************************************************************
Public Property Get RemotePort() As Long

  Dim SockAddr As API_SOCKADDR_IN
  Dim ErrorObject As CWinsockException

    If m_lngSocketHandle = INVALID_SOCKET Then Exit Property

    If api_GetPeerName(m_lngSocketHandle, SockAddr, LenB(SockAddr)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.RemotePort"
        RaiseEvent OnError(ErrorObject)
    Else
        RemotePort = IntegerToUnsigned(api_ntohs(SockAddr.sin_port))
    End If

End Property


' *************************************************************************************************
' Description:  Open the socket.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     03/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function OpenSocket(Optional ByVal AddressFamily As AddressFamilyType = AddFam_InterNetwork, _
                           Optional ByVal SockType As SocketType = SockType_stream, _
                           Optional ByVal Protocol As IPProtocolType = Proto_Tcp, _
                           Optional ByVal Async As Boolean = False) As Boolean

    
  Dim ErrorObject As CWinsockException
  
    ' Make sure the socket is closed before trying to open it again
    If m_lngSocketHandle <> INVALID_SOCKET Then
        Err.Raise -1, "CSocket.OpenSocket", "The socket is already opened. Close it first before trying to open it again"
        Exit Function
    End If
    
    ' Attempt to create the socket
    m_lngSocketHandle = api_Socket(AddressFamily, SockType, Protocol)

    ' Initialize socket
    OpenSocket = InitializeSocket(AddressFamily, SockType, Protocol, Async)

End Function


' *************************************************************************************************
' Description:  Initialize a socket
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     20/04/2004 by Chris Waddell
' *************************************************************************************************
Friend Function InitializeSocket(Optional ByVal AddressFamily As AddressFamilyType = AddFam_InterNetwork, _
                                 Optional ByVal SockType As SocketType = SockType_stream, _
                                 Optional ByVal Protocol As IPProtocolType = Proto_Tcp, _
                                 Optional ByVal Async As Boolean = False) As Boolean

    
  Dim ErrorObject As CWinsockException
  
    ' Check for an error
    If m_lngSocketHandle = INVALID_SOCKET Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.InitializeSocket"
        RaiseEvent OnError(ErrorObject)
    Else
        m_AddressFamily = AddressFamily
        m_SocketType = SockType
        m_Protocol = Protocol
    
        Set m_SocketIO = New CSocketIO
        Set m_SocketOptions = New CSocketOptions
        
        m_SocketIO.SocketHandle = m_lngSocketHandle
        m_SocketOptions.SocketHandle = m_lngSocketHandle
        
        SetBlocking Async
        
        InitializeSocket = True
    End If

End Function


' *************************************************************************************************
' Description:  Shutdown the socket for a graceful close.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     03/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function ShutDown(ShutdownType As SocketShutdownType)

  Dim ErrorObject As CWinsockException

    If m_lngSocketHandle = INVALID_SOCKET Then Exit Function

    If api_Shutdown(m_lngSocketHandle, ShutdownType) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.Shutdown"
        RaiseEvent OnError(ErrorObject)
    End If
    
End Function


' *************************************************************************************************
' Description:  Close the socket.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     03/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function CloseSocket() As Boolean

  Dim RetVal As Long
  Dim ErrorObject As CWinsockException

    ' Make sure the socket is actually opened before trying to close it
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.OpenSocket", "The socket is already closed."
        Exit Function
    End If
    
    ' Attempt to close the socket
    RetVal = api_CloseSocket(m_lngSocketHandle)
    
    ' Check for an error
    If RetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.CloseSocket"
        RaiseEvent OnError(ErrorObject)
    Else
        Set m_SocketIO = Nothing
        Set m_SocketOptions = Nothing
    
        m_lngSocketHandle = INVALID_SOCKET
        CloseSocket = True
    End If

End Function


' *************************************************************************************************
' Description:  Bind the socket to an address. Returns true if successful and false if
'               otherwise. If it returns false, then the OnError event will be raised with specific
'               information about the nature of the error.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Bind(Address As CIP4Address, Port As Long) As Boolean

  Dim SockAddr      As API_SOCKADDR_IN
  Dim ErrorObject   As CWinsockException
  

    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.Bind", "The socket handle hasn't been created yet."
        Exit Function
    End If
  
    ' Check to make sure the address is valid
    If Address Is Nothing Then
        Err.Raise -1, "CSocket.Bind", "There is no address."
        Exit Function
    End If

    ' Check to make sure the port is valid
    If Port < 0 Or Port > 65535 Then
        Err.Raise -1, "CSocket.Bind", "The port is invalid."
        Exit Function
    End If
  
    ' Fill out the sockaddr structure
    SockAddr.sin_addr.S_addr = api_htonl(Address.Address)
    SockAddr.sin_family = m_AddressFamily
    SockAddr.sin_port = api_htons(UnsignedToInteger(Port))

    ' Attempt to bind the socket
    If api_Bind(m_lngSocketHandle, SockAddr, LenB(SockAddr)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.Bind"
        RaiseEvent OnError(ErrorObject)
    Else
        Bind = True
    End If

End Function


' *************************************************************************************************
' Description:  Listens for connections. Returns true if successful and false if otherwise. If it
'               returns false, then the OnError event will be raised with specific information about
'               the nature of the error.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Listen(Optional Backlog As Long = SOMAXCONN) As Boolean

  Dim ErrorObject   As CWinsockException

    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.Listen", "The socket handle hasn't been created yet."
        Exit Function
    End If

    ' Attempt to set the socket into listening mode
    If api_Listen(m_lngSocketHandle, Backlog) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.Listen"
        RaiseEvent OnError(ErrorObject)
    Else
        Listen = True
    End If

End Function


' *************************************************************************************************
' Description:  Attempt to connect to the remote host. Returns true if successful and false if
'               otherwise. If it returns false, then the OnError event will be raised with specific
'               information about the nature of the error.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     09/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Connect(Address As CIP4Address, Port As Long) As Boolean

  Dim SockAddr      As API_SOCKADDR_IN
  Dim ErrorObject   As CWinsockException
  
  
    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.Connect", "The socket handle hasn't been created yet."
        Exit Function
    End If
  
    ' Check to make sure the address is valid
    If Address Is Nothing Then
        Err.Raise -1, "CSocket.Connect", "There is no address."
        Exit Function
    End If
    
    ' Check to make sure the port is valid
    If Port < 0 Or Port > 65535 Then
        Err.Raise -1, "CSocket.Connect", "The port is invalid."
        Exit Function
    End If
  
    ' Fill out the sockaddr structure
    SockAddr.sin_addr.S_addr = Address.Address
    SockAddr.sin_family = m_AddressFamily
    SockAddr.sin_port = api_htons(UnsignedToInteger(Port))

    ' Attempt to connect
    If api_Connect(m_lngSocketHandle, SockAddr, LenB(SockAddr)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.Connect"
        RaiseEvent OnError(ErrorObject)
    Else
        Connect = True
    End If

End Function


' *************************************************************************************************
' Description:  Listens for connections. Returns true if successful and false if otherwise. If it
'               returns false, then the OnError event will be raised with specific information about
'               the nature of the error.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Accept(Address As CIP4Address, RemotePort As Long) As CSocket

  Dim ErrorObject   As CWinsockException
  Dim lngRetVal     As Long
  Dim SockAddr      As API_SOCKADDR_IN

    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.Accept", "The socket handle hasn't been created yet."
        Exit Function
    End If
    
    ' Try to accept the connection.
    lngRetVal = api_Accept(m_lngSocketHandle, SockAddr, LenB(SockAddr))
    
    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.Accept"
        RaiseEvent OnError(ErrorObject)
    Else
        Set Address = New CIP4Address
        Address.Address = api_ntohl(SockAddr.sin_addr.S_addr)
        RemotePort = api_ntohs(SockAddr.sin_port)
    
        Set Accept = New CSocket
        Accept.SocketHandle = lngRetVal
        Accept.InitializeSocket m_AddressFamily, m_SocketType, m_Protocol, m_Async
    End If

End Function


' *************************************************************************************************
' Description:  Waits for the socket to enter a writable state. The time out is in milliseconds.
'               If it is 0, it will have no effect. Returns true if the socket is now in a writable
'               state or false if it timed out or encountered an error
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function WaitForWritable(Optional TimeoutTime As Long = 0) As Boolean

  Dim WriteInfo     As API_FD_SET
  Dim TimeOut       As API_TIMEVAL
  Dim lngRetVal     As Long
  Dim ErrorObject   As CWinsockException

    WriteInfo.fd_array(0) = m_lngSocketHandle
    WriteInfo.fd_count = 1
    
    If TimeoutTime > 0 Then
    
        ' Convert to seconds and microseconds
        TimeOut.tv_sec = TimeoutTime / 1000
        TimeOut.tv_usec = (TimeoutTime - (TimeOut.tv_sec * 1000)) * 1000

        lngRetVal = api_Select(ByVal 0&, WriteInfo, ByVal 0&, ByVal 0&, ByVal VarPtr(TimeOut))
    Else
        lngRetVal = api_Select(ByVal 0&, WriteInfo, ByVal 0&, ByVal 0&, ByVal 0&)
    End If


    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.WaitForWritable"
        RaiseEvent OnError(ErrorObject)
    Else
        WaitForWritable = CBool(lngRetVal)
    End If

End Function


' *************************************************************************************************
' Description:  Waits for the socket to enter a readable state. The time out is in milliseconds.
'               If it is 0, it will have no effect. Returns true if the socket is now in a readable
'               state or false if it timed out or encountered an error
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function WaitForReadable(Optional TimeoutTime As Long = 0) As Boolean

  Dim ReadInfo      As API_FD_SET
  Dim TimeOut       As API_TIMEVAL
  Dim lngRetVal     As Long
  Dim ErrorObject   As CWinsockException

    ReadInfo.fd_array(0) = m_lngSocketHandle
    ReadInfo.fd_count = 1
    
    If TimeoutTime > 0 Then
    
        ' Convert to seconds and microseconds
        TimeOut.tv_sec = TimeoutTime / 1000
        TimeOut.tv_usec = (TimeoutTime - (TimeOut.tv_sec * 1000)) * 1000

        lngRetVal = api_Select(ByVal 0&, ByVal 0&, ReadInfo, ByVal 0&, ByVal VarPtr(TimeOut))
    Else
        lngRetVal = api_Select(ByVal 0&, ByVal 0&, ReadInfo, ByVal 0&, ByVal 0&)
    End If


    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.WaitForReadable"
        RaiseEvent OnError(ErrorObject)
    Else
        WaitForReadable = CBool(lngRetVal)
    End If

End Function


' *************************************************************************************************
' Description:  Recieve data from a connected stream oriented socket and return the number of bytes
'               read.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Recieve(Buffer() As Byte, Optional Flags As MessageCommunicationFlags = Msg_None) As Long

  Dim RecieveBuffer()   As Byte
  Dim ErrorObject       As CWinsockException
  Dim lngRetVal         As Long


    If Not m_SocketIO Is Nothing Then
        ReDim RecieveBuffer(0 To m_SocketIO.GetReadLength - 1)
    Else
        ReDim RecieveBuffer(0 To 1023)
    End If

    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.Recieve", "The socket handle hasn't been created yet."
        Exit Function
    End If

    ' Attempt to recieve some data
    lngRetVal = api_Recv(m_lngSocketHandle, RecieveBuffer(0), UBound(RecieveBuffer) + 1, Flags)

    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.Recieve"
        RaiseEvent OnError(ErrorObject)
    Else
        ' Resize the buffer
        ReDim Buffer(0 To lngRetVal - 1)
        
        ' Copy the recieved data into it
        RtlMoveMemory ByVal VarPtr(Buffer(0)), ByVal VarPtr(RecieveBuffer(0)), lngRetVal
        
        ' Return the number of bytes recieved
        Recieve = lngRetVal
    End If

End Function


' *************************************************************************************************
' Description:  Recieve data from a remote host using a connectionless socket and return the number
'               of bytes read.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function RecieveFrom(Buffer() As Byte, RemoteHost As CIP4Address, RemotePort As Long, Optional Flags As MessageCommunicationFlags = Msg_None) As Long

  Dim RecieveBuffer()   As Byte
  Dim ErrorObject       As CWinsockException
  Dim lngRetVal         As Long
  Dim SockAddr          As API_SOCKADDR_IN


    If Not m_SocketOptions Is Nothing Then
        ReDim RecieveBuffer(0 To m_SocketOptions.SocketOption_MaxMsgSize - 1)
    Else
        ReDim RecieveBuffer(0 To 1023)
    End If

    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.RecieveFrom", "The socket handle hasn't been created yet."
        Exit Function
    End If

    ' Fill out the sockaddr structure
    SockAddr.sin_addr.S_addr = RemoteHost.Address
    SockAddr.sin_family = m_AddressFamily
    SockAddr.sin_port = api_htons(UnsignedToInteger(RemotePort))

    ' Attempt to recieve some data
    lngRetVal = api_RecvFrom(m_lngSocketHandle, RecieveBuffer(0), UBound(RecieveBuffer) + 1, Flags, SockAddr, LenB(SockAddr))

    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.RecieveFrom"
        RaiseEvent OnError(ErrorObject)
    Else
        ' Resize the buffer
        ReDim Buffer(0 To lngRetVal - 1)
        
        ' Copy the recieved data into it
        RtlMoveMemory ByVal VarPtr(Buffer(0)), ByVal VarPtr(RecieveBuffer(0)), lngRetVal
        
        ' Return the number of bytes recieved
        RecieveFrom = lngRetVal
    End If

End Function


' *************************************************************************************************
' Description:  Send data to a remote host using a connectionless socket and return the number
'               of bytes sent.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function SendTo(Buffer() As Byte, RemoteHost As CIP4Address, RemotePort As Long, Optional Flags As MessageCommunicationFlags = Msg_None) As Long

On Error GoTo ErrHandler

  Dim ErrorObject       As CWinsockException
  Dim lngRetVal         As Long
  Dim SockAddr          As API_SOCKADDR_IN

    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.SendTo", "The socket handle hasn't been created yet."
        Exit Function
    End If

    ' Fill out the sockaddr structure
    SockAddr.sin_addr.S_addr = RemoteHost.Address
    SockAddr.sin_family = m_AddressFamily
    SockAddr.sin_port = api_htons(UnsignedToInteger(RemotePort))

    ' Attempt to send some data
    lngRetVal = api_SendTo(m_lngSocketHandle, Buffer(0), UBound(Buffer) + 1, Flags, SockAddr, LenB(SockAddr))

    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.SendTo"
        RaiseEvent OnError(ErrorObject)
    Else
        ' Return the number of bytes sent
        SendTo = lngRetVal
    End If

    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CSocket.SendTo", Err.Description

End Function


' *************************************************************************************************
' Description:  Sends data through the socket and returns the number of bytes sent.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Send(Buffer() As Byte, Optional Flags As MessageCommunicationFlags = Msg_None) As Long

On Error GoTo ErrHandler

  Dim ErrorObject       As CWinsockException
  Dim lngRetVal         As Long

    ' Check to make sure the socket handle is valid
    If m_lngSocketHandle = INVALID_SOCKET Then
        Err.Raise -1, "CSocket.Send", "The socket handle hasn't been created yet."
        Exit Function
    End If

    ' Attempt to send some data
    lngRetVal = api_Send(m_lngSocketHandle, Buffer(0), UBound(Buffer) + 1, Flags)

    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocket.Send"
        RaiseEvent OnError(ErrorObject)
    Else
        ' Return the number of bytes sent
        Send = lngRetVal
    End If

    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CSocket.Send", Err.Description

End Function



Private Sub EventWindow_MessageArrived(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, lngRetVal As Long, bHandled As Boolean)

  Dim lngErrorCode          As Long
  Dim lngEvent              As Long
  Dim lngBytesRecieved      As Long
  Dim ErrorObject           As CWinsockException
  Dim Buffer()              As Byte

    ' Extract the error code and event from the lParam
    lngErrorCode = HiWord(lParam)
    lngEvent = LoWord(lParam)
    
    ' Check for errors
    If lngErrorCode <> 0 Then
        Set ErrorObject = New CWinsockException
        ErrorObject.ErrorCode = lngErrorCode
        ErrorObject.Source = "CSocket.EventWindow_MessageArrived"
        RaiseEvent OnError(ErrorObject)
        Exit Sub
    End If
    
    ' Check what the event is
    Select Case lngEvent
        Case AsyncSelectFlagsType.ASFlag_Accept
            RaiseEvent ConnectionRequest
        Case AsyncSelectFlagsType.ASFlag_Close
            CloseSocket
            RaiseEvent Closed
        Case AsyncSelectFlagsType.ASFlag_Connect
            RaiseEvent Connected
        Case AsyncSelectFlagsType.ASFlag_Read
            RaiseEvent DataArrived
        Case AsyncSelectFlagsType.ASFlag_Write
            RaiseEvent ReadyForWriting
        Case Else
            Debug.Print "Unknown event"
    End Select

    ' We process this one
    lngRetVal = 0
    bHandled = True

End Sub


Private Sub m_SocketIO_OnError(Exception As CWinsockException)
    RaiseEvent OnError(Exception)
End Sub

Private Sub m_SocketOptions_OnError(Exception As CWinsockException)
    RaiseEvent OnError(Exception)
End Sub


Private Sub Class_Initialize()
    m_lngSocketHandle = INVALID_SOCKET
    Set EventWindow = New CWindow
End Sub


Private Sub Class_Terminate()

    ' Close the socket if it's not already closed
    If m_lngSocketHandle <> INVALID_SOCKET Then
        CloseSocket
    End If
    
    Set EventWindow = Nothing
    
End Sub
