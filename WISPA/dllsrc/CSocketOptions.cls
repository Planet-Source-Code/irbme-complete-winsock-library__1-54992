VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocketOptions"
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
    SocketHandle = m_lngSocketHandle
End Property

Public Property Let SocketHandle(Value As Long)
    m_lngSocketHandle = Value
End Property


' *************************************************************************************************
' Socket is listening.
' *************************************************************************************************
Public Property Get SocketOption_Listening() As Boolean
    SocketOption_Listening = CBool(GetOptionLong(SOL_SOCKET, SO_ACCEPTCONN, "Get SocketOption_Listening"))
End Property


' *************************************************************************************************
' Socket is configured for the transmission of broadcast messages.
' *************************************************************************************************
Public Property Get SocketOption_Broadcast() As Boolean
    SocketOption_Broadcast = CBool(GetOptionLong(SOL_SOCKET, SO_BROADCAST, "Get SocketOption_Broadcast"))
End Property

Public Property Let SocketOption_Broadcast(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_BROADCAST, CLng(Value), "Let SocketOption_Broadcast"
End Property



' *************************************************************************************************
' Debugging is enabled.
' *************************************************************************************************
Public Property Get SocketOption_DebugEnabled() As Boolean
    SocketOption_DebugEnabled = CBool(GetOptionLong(SOL_SOCKET, SO_DEBUG, "Get SocketOption_DebugEnabled"))
End Property

Public Property Let SocketOption_DebugEnabled(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_DEBUG, CLng(Value), "Let SocketOption_DebugEnabled"
End Property


' *************************************************************************************************
' If TRUE, the SO_LINGER option is disabled.
' *************************************************************************************************
Public Property Get SocketOption_DontLinger() As Boolean
    SocketOption_DontLinger = CBool(GetOptionLong(SOL_SOCKET, SO_DONTLINGER, "Get SocketOption_DontLinger"))
End Property

Public Property Let SocketOption_DontLinger(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_DONTLINGER, CLng(Value), "Let SocketOption_DontLinger"
End Property


' *************************************************************************************************
' Routing is disabled. Succeeds but is ignored on AF_INET sockets; fails on AF_INET6
' sockets with WSAENOPROTOOPT. Not supported on ATM sockets.
' *************************************************************************************************
Public Property Get SocketOption_DontRoute() As Boolean
    SocketOption_DontRoute = CBool(GetOptionLong(SOL_SOCKET, SO_DONTROUTE, "Get SocketOption_DontRoute"))
End Property

Public Property Let SocketOption_DontRoute(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_DONTROUTE, CLng(Value), "Let SocketOption_DontRoute"
End Property

' *************************************************************************************************
' Retrieves error status and clear.
' *************************************************************************************************

Public Property Get SocketOption_ErrorStatus() As Long
    SocketOption_ErrorStatus = GetOptionLong(SOL_SOCKET, SO_ERROR, "Get SocketOption_ErrorStatus")
End Property


' *************************************************************************************************
' Keep-alives are being sent. Not supported on ATM sockets.
' *************************************************************************************************
Public Property Get SocketOption_KeepAlive() As Boolean
    SocketOption_KeepAlive = CBool(GetOptionLong(SOL_SOCKET, SO_KEEPALIVE, "Get SocketOption_KeepAlive"))
End Property

Public Property Let SocketOption_KeepAlive(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_DONTROUTE, CLng(Value), "Let SocketOption_KeepAlive"
End Property


' *************************************************************************************************
' Returns the current linger options.
' *************************************************************************************************
Public Property Get SocketOption_Linger() As CLinger

  Dim Linger As API_LINGER
  Dim ErrorObject As CWinsockException

    If api_GetSockOpt(m_lngSocketHandle, SOL_SOCKET, SO_LINGER, Linger, LenB(Linger)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocketOptions.Get SocketOption_Linger"
                
        RaiseEvent OnError(ErrorObject)
    Else
        Set SocketOption_Linger = New CLinger
        Call SocketOption_Linger.SetLinger(VarPtr(Linger))
    End If

End Property

Public Property Set SocketOption_Linger(Value As CLinger)

  Dim Linger As API_LINGER
  Dim ErrorObject As CWinsockException
  
    If Value Is Nothing Then Exit Property
    
    Linger.l_onoff = IIf(Value.LingerEnabled, 1, 0)
    Linger.l_linger = Value.TimeOut
    
    If api_SetSockOpt(m_lngSocketHandle, SOL_SOCKET, SO_LINGER, Linger, LenB(Linger)) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocketOptions.Set SocketOption_Linger"
                
        RaiseEvent OnError(ErrorObject)
    End If

End Property


' *************************************************************************************************
' Maximum size of a message for message-oriented socket types (for example,
' SOCK_DGRAM). Has no meaning for stream oriented sockets.
' *************************************************************************************************
Public Property Get SocketOption_MaxMsgSize() As Currency
    SocketOption_MaxMsgSize = LongToUnsigned(GetOptionLong(SOL_SOCKET, SO_MAX_MSG_SIZE, "Get SocketOption_MaxMsgSize"))
End Property


' *************************************************************************************************
' OOB data is being received in the normal data stream. (See section Windows Sockets
' 1.1 Blocking Routines and EINPROGRESS for a discussion of this topic.)
' *************************************************************************************************
Public Property Get SocketOption_OOBInline() As Boolean
    SocketOption_OOBInline = GetOptionLong(SOL_SOCKET, SO_OOBINLINE, "Get SocketOption_OOBInline")
End Property

Public Property Let SocketOption_OOBInline(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_OOBINLINE, CLng(Value), "Let SocketOption_OOBInline"
End Property


' *************************************************************************************************
' Description of protocol information for protocol that is bound to this socket.
' *************************************************************************************************
Public Property Get SocketOption_ProtocolInfo() As CProtocolInfo
    Set m_ProtocolInfo = New CProtocolInfo
    m_ProtocolInfo.GetProtocolByHandle (m_lngSocketHandle)
    Set SocketOption_ProtocolInfo = m_ProtocolInfo
End Property


' *************************************************************************************************
' Buffer size for receives.
' *************************************************************************************************
Public Property Get SocketOption_RecvBufferSize() As Currency
    SocketOption_RecvBufferSize = GetOptionLong(SOL_SOCKET, SO_RCVBUF, "Get SocketOption_RecvBufferSize")
End Property


' *************************************************************************************************
' The socket can be bound to an address which is already in use. Not applicable for ATM sockets.
' *************************************************************************************************
Public Property Get SocketOption_ReuseAddress() As Boolean
    SocketOption_ReuseAddress = CBool(GetOptionLong(SOL_SOCKET, SO_REUSEADDR, "Get SocketOption_ReuseAddress"))
End Property

Public Property Let SocketOption_ReuseAddress(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_REUSEADDR, CLng(Value), "Let SocketOption_ReuseAddress"
End Property

Public Property Let SocketOption_ExclusiveUse(Value As Boolean)
    SetOptionLong SOL_SOCKET, SO_EXCLUSIVEADDRUSE, CLng(Value), "Let SocketOption_ExclusiveUse"
End Property


' *************************************************************************************************
' Buffer size for sends.
' *************************************************************************************************
Public Property Get SocketOption_SendBufferSize() As Currency
    SocketOption_SendBufferSize = GetOptionLong(SOL_SOCKET, SO_SNDBUF, "Get SocketOption_SendBufferSize")
End Property

Public Property Let SocketOption_SendBufferSize(Value As Currency)
    SetOptionLong SOL_SOCKET, SO_SNDBUF, CLng(UnsignedToLong(Value)), "Let SocketOption_SendBufferSize"
End Property

' *************************************************************************************************
' The type of the socket (for example, SOCK_STREAM).
' *************************************************************************************************
Public Property Get SocketOption_SocketType() As SocketType
    SocketOption_SocketType = GetOptionLong(SOL_SOCKET, SO_TYPE, "Get SocketOption_SocketType")
End Property


' *************************************************************************************************
' Simple function which removes the need for all of the above properties to do error checking.
' This function only works for options which return 4 bytes of data such as "int"s or "BOOL"s.
' *************************************************************************************************
Private Function GetOptionLong(Level As SocketOptionLevel, Name As SocketOptionName, Source As String) As Long
    
  Dim ErrorObject As CWinsockException
  
    If api_GetSockOpt(m_lngSocketHandle, Level, Name, GetOptionLong, 4) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocketOptions." & Source
                
        RaiseEvent OnError(ErrorObject)
    End If
End Function


' *************************************************************************************************
' Simple sub which removes the need for all of the above properties to do error checking.
' This function only works for options which have 4 bytes of data such as "int"s or "BOOL"s.
' *************************************************************************************************
Private Sub SetOptionLong(Level As SocketOptionLevel, Name As SocketOptionName, Value As Long, Source As String)
    
  Dim ErrorObject As CWinsockException
  
    If api_SetSockOpt(m_lngSocketHandle, Level, Name, Value, 4) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CSocketOptions." & Source
                
        RaiseEvent OnError(ErrorObject)
    End If
End Sub


' *************************************************************************************************
' When the protocol info class encounters an error, move it up the stack to the caller of this
' class.
' *************************************************************************************************
Private Sub m_ProtocolInfo_OnError(Exception As CWinsockException)
    RaiseEvent OnError(Exception)
End Sub
