VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocketIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Exception As CWinsockException)


' The local copy of the socket handle
Private m_lngSocketHandle As Long

' An instance of the protocol info class
Private WithEvents m_ProtocolInfo As CProtocolInfo
Attribute m_ProtocolInfo.VB_VarHelpID = -1


' *************************************************************************************************
' The socket handle of the socket we want to get info about.
' *************************************************************************************************
Public Property Get SocketHandle() As Long
Attribute SocketHandle.VB_MemberFlags = "40"
    SocketHandle = m_lngSocketHandle
End Property

Public Property Let SocketHandle(Value As Long)
    m_lngSocketHandle = Value
End Property


' *************************************************************************************************
' Enables or disables nonblocking mode on socket s. lpvInBuffer points at an unsigned long, which
' is nonzero if nonblocking mode is to be enabled and zero if it is to be disabled. When a socket
' is created, it operates in blocking mode (that is, nonblocking mode is disabled). This is
' consistent with Berkeley Software Distribution (BSD) sockets.
' *************************************************************************************************
Public Property Let SetBlockingMode(Value As Boolean)

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    ' Set to 0 for blocking mode, or non-zero for non blocking mode
    InBuffer = IIf(Value, 0, 1)
  
    If api_WSAIoctl(m_lngSocketHandle, FIONBIO, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0, 0&) Then
        RaiseError "BlockingMode"
    End If
    
End Property


' *************************************************************************************************
' Determines the amount of data that can be read atomically from socket s. lpvOutBuffer points at
' an unsigned long in which WSPIoctl stores the result. If s is stream oriented (for example,
' type SOCK_STREAM), FIONREAD returns the total amount of data that can be read in a single
' receive operation; this is normally the same as the total amount of data queued on the socket.
' If s is message oriented (for example, type SOCK_DGRAM), FIONREAD returns the size of the first
' datagram (message) queued on the socket.
' *************************************************************************************************
Public Property Get GetReadLength() As Currency

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    If api_WSAIoctl(m_lngSocketHandle, FIONREAD, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "ReadLength"
    Else
        GetReadLength = LongToUnsigned(OutBuffer)
    End If

End Property


' *************************************************************************************************
' Determines whether or not all OOB data has been read. This applies only to a socket of stream
' style (for example, type SOCK_STREAM) that has been configured for inline reception of any OOB
' data (SO_OOBINLINE). If no OOB data is waiting to be read, the operation returns TRUE.
' Otherwise, it returns FALSE, and the next receive operation performed on the socket will
' retrieve some or all of the data preceding the mark; the Windows Sockets SPI client should use
' the SIOCATMARK operation to determine whether any remains. If there is any normal data
' preceding the urgent (OOB) data, it will be received in order. (Note that receive operations
' will never mix OOB and normal data in the same call.) lpvOutBuffer points at a bool in which
' WSPIoctl stores the result.
' *************************************************************************************************
Public Property Get IsUrgentInlineDataWaiting() As Boolean

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    If api_WSAIoctl(m_lngSocketHandle, SIOCATMARK, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "UrgentInlineDataWaiting"
    Else
        IsUrgentInlineDataWaiting = CBool(OutBuffer)
    End If

End Property


' *************************************************************************************************
' Enables a socket to receive all IP packets on the network. The socket handle passed to the
' WSAIoctl function must be of AF_INET address family, SOCK_RAW socket type, and IPPROTO_IP
' protocol. The socket also must be bound to an explicit local interface, which means that you
' cannot bind to INADDR_ANY.
'
' Once the socket is bound and the ioctl set, calls to the WSARecv or recv functions return IP
' datagrams passing through the given interface. Note that you must supply a sufficiently large
' buffer. Setting this ioctl requires Administrator privilege on the local computer. SIO_RCVALL
' is available in Windows 2000 and later versions of Windows.
' *************************************************************************************************
Public Property Let RecieveAll(Value As Boolean)

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    ' 1 for enabled, 0 for disabled
    InBuffer = IIf(Value, 1, 0)
  
    If api_WSAIoctl(m_lngSocketHandle, SIO_RCVALL, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "RecieveAll"
    End If

End Property


' *************************************************************************************************
' Enables a socket to receive all multicast IP traffic on the network (that is, all IP packets
' destined for IP addresses in the range of 224.0.0.0 to 239.255.255.255). The socket handle
' passed to the WSAIoctl function must be of AF_INET address family, SOCK_RAW socket type, and
' IPPROTO_UDP protocol. The socket also must bind to an explicit local interface, which means
' that you cannot bind to INADDR_ANY. The socket should bind to port zero.

' Once the socket is bound and the ioctl set, calls to the WSARecv or recv functions return
' multicast IP datagrams passing through the given interface. Note that you must supply a
' sufficiently large buffer. Setting this ioctl requires Administrator privilege on the local
' computer. SIO_RCVALL_MCAST is available only in Windows 2000 and later versions of Windows.
' *************************************************************************************************
Public Property Let RecieveAllMulticast(Value As Boolean)

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    ' 1 for enabled, 0 for disabled
    InBuffer = IIf(Value, 1, 0)
  
    If api_WSAIoctl(m_lngSocketHandle, SIO_RCVALL_MCAST, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "RecieveAllMulticast"
    End If

End Property


' *************************************************************************************************
' Enables a socket to receive all IGMP multicast IP traffic on the network, without receiving
' other multicast IP traffic. The socket handle passed to the WSAIoctl function must be of
' AF_INET address family, SOCK_RAW socket type, and IPPROTO_IGMP protocol. The socket also must
' be bound to an explicit local interface, which means that you cannot bind to INADDR_ANY.
'
' Once the socket is bound and the ioctl set, calls to the WSARecv or recv functions return
' multicast IP datagrams passing through the given interface. Note that you must supply a
' sufficiently large buffer. Setting this ioctl requires Administrator privilege on the local
' computer. SIO_RCVALL_IGMPMCAST is available only in Windows 2000 and later versions of Windows.
' *************************************************************************************************
Public Property Let RecieveAllIGMPcast(Value As Boolean)

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    ' 1 for enabled, 0 for disabled
    InBuffer = IIf(Value, 1, 0)
  
    If api_WSAIoctl(m_lngSocketHandle, SIO_RCVALL_IGMPMCAST, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "RecieveAllIGMPcast"
    End If

End Property


' *************************************************************************************************
' Indicates to a message-oriented service provider that a newly arrived message should never be
' dropped because of a buffer queue overflow. Instead, the oldest message in the queue should be
' eliminated in order to accommodate the newly arrived message. No input and output buffers are
' required. Note that this IOCTL is only valid for sockets associated with unreliable,
' message-oriented protocols.
' *************************************************************************************************
Public Sub EnableCircularQueueing()
  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long

    If api_WSAIoctl(m_lngSocketHandle, SIO_ENABLE_CIRCULAR_QUEUEING, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "EnableCircularQueueing"
    End If
End Sub


' *************************************************************************************************
' This IOCTL fills the output buffer with a sockaddr structure containing a suitable broadcast
' address for use with WSASendTo.
' *************************************************************************************************
Public Property Get GetSuitableBroadcastAddress() As CSocketAddress

  Dim InBuffer As Long, OutBuffer As API_SOCKADDR_IN, BytesReturned As Long, lngRetVal As Long
    
    lngRetVal = api_WSAIoctl(m_lngSocketHandle, SIO_GET_BROADCAST_ADDRESS, ByVal InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, ByVal 0&, ByVal 0&)
    
    If lngRetVal = SOCKET_ERROR Then
        RaiseError "SuitableBroadcastAddress"
    Else
        Set GetSuitableBroadcastAddress = New CSocketAddress
        GetSuitableBroadcastAddress.GetBySockAddr VarPtr(OutBuffer)
    End If

End Property


' *************************************************************************************************
' Controls whether data sent in a multipoint session will also be received by the same socket on
' the local host. A value of TRUE causes loopback reception to occur while a value of FALSE
' prohibits this.
' *************************************************************************************************
Public Property Let EnableMultiPointLoopBack(Value As Boolean)
  
  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    ' 1 for enabled, 0 for disabled
    InBuffer = IIf(Value, 1, 0)
  
    If api_WSAIoctl(m_lngSocketHandle, SIO_MULTIPOINT_LOOPBACK, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "EnableMultiPointLoopBack"
    End If
    
End Property


' *************************************************************************************************
' Specifies the scope over which multicast transmissions will occur. Scope is defined as the
' number of routed network segments to be covered. A scope of zero would indicate that the
' multicast transmission would not be placed on the wire, but could be disseminated across
' sockets within the local host. A scope value of 1 (the default) indicates that the transmission
' will be placed on the wire, but will not cross any routers. Higher scope values determine the
' number of routers that can be crossed. Note that this corresponds to the time-to-live (TTL)
' parameter in IP multicasting.
' *************************************************************************************************
Public Property Let SetMulticastScope(Value As Long)

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long
  
    InBuffer = Value
  
    If api_WSAIoctl(m_lngSocketHandle, SIO_MULTICAST_SCOPE, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "MulticastScope"
    End If
    
End Property


' *************************************************************************************************
' Discards current contents of the sending queue associated with this socket. No input and output
' buffers are required. The WSAENOPROTOOPT error code is indicated for service providers that do
' not support this IOCTL.
' *************************************************************************************************
Public Sub FlushSendQueue()

  Dim InBuffer As Long, OutBuffer As Long, BytesReturned As Long

    If api_WSAIoctl(m_lngSocketHandle, SIO_FLUSH, InBuffer, LenB(InBuffer), OutBuffer, LenB(OutBuffer), BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        RaiseError "FlushSendQueue"
    End If
    
End Sub


' *************************************************************************************************
' Raise the last dll error which occured.
' *************************************************************************************************
Private Function RaiseError(Source As String)

  Dim ErrorObject As CWinsockException
    Set ErrorObject = New CWinsockException
    
    ErrorObject.Source = "CSocketIO." & Source
    RaiseEvent OnError(ErrorObject)
    
End Function
