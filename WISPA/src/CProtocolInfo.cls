VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProtocolInfo"
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


' A private copy of the protocol info structure from which properties can read the data from
Private m_ProtocolInfo As API_WSAPROTOCOL_INFO


' *************************************************************************************************
' Description: Get the protocol information from a socket handle. The protocol type will be the
'              protocol used by the socket
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    04/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub GetProtocolByHandle(lngSocketHandle As Long)

  Dim lngBuffLen As Long
  Dim ErrorObject As CWinsockException
      
    lngBuffLen = LenB(m_ProtocolInfo)

    If api_GetSockOpt(lngSocketHandle, SOL_SOCKET, SO_PROTOCOL_INFO, m_ProtocolInfo, lngBuffLen) = SOCKET_ERROR Then
    
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CProtocolInfo.GetProtocolByHandle"
                
        RaiseEvent OnError(ErrorObject)
    End If
    
End Sub


' *************************************************************************************************
' Description: Set the protocol info. lpProtocolInfo MUST be a valid pointer to a
'              API_PROTOCOL_INFO structure or else a general protection fault is likely to occur
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    04/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub GetProtocolByProtocolInfoType(lpProtocolInfo As Long)
    If lpProtocolInfo <= 0 Then Exit Sub
    RtlMoveMemory ByVal VarPtr(m_ProtocolInfo), lpProtocolInfo, LenB(m_ProtocolInfo)
End Sub



' *************************************************************************************************
' The name of the protocol.
' *************************************************************************************************
Public Property Get ProtocolName() As String
    ProtocolName = Replace$(m_ProtocolInfo.szProtocol, Chr$(0), vbNullString)
End Property


' *************************************************************************************************
' Maximum message size supported by the protocol. This is the maximum size, in bytes, that can be
' sent from any of the host's local interfaces. For protocols that do not support message framing,
' the actual maximum that can be sent to a given address may be less. There is no standard
' provision to determine the maximum inbound message size.
' *************************************************************************************************
Public Property Get MessageSize() As Long
    MessageSize = m_ProtocolInfo.dwMessageSize
End Property


' *************************************************************************************************
' Indicates the type of security scheme employed (if any). A value of SECURITY_PROTOCOL_NONE is
' used for protocols that do not incorporate security provisions.
' *************************************************************************************************
Public Property Get SecurityScheme() As Long
    SecurityScheme = m_ProtocolInfo.iSecurityScheme
End Property


' *************************************************************************************************
' Currently these values are manifest constants (BIGENDIAN and LITTLEENDIAN) that indicate either
' big-endian or little-endian with the values 0 and 1 respectively.
' *************************************************************************************************
Public Property Get NetworkByteOrder() As ByteOrder
    NetworkByteOrder = m_ProtocolInfo.iNetworkByteOrder
End Property


' *************************************************************************************************
' Maximum value that may be added to iProtocol when supplying a value for the protocol parameter
' to socket and WSASocket. Not all protocols allow a range of values. When this is the case
' iProtocolMaxOffset is zero.
' *************************************************************************************************
Public Property Get ProtocolMaxOffset() As Long
    ProtocolMaxOffset = m_ProtocolInfo.iProtocolMaxOffset
End Property


' *************************************************************************************************
' Value to pass as the protocol parameter to the socket function in order to open a socket for
' this protocol.
' *************************************************************************************************
Public Property Get Protocol() As Long
    Protocol = m_ProtocolInfo.iProtocol
End Property


' *************************************************************************************************
' Value to pass as the socket type parameter to the socket function in order to open a socket for
' this protocol.
' *************************************************************************************************
Public Property Get SocketType() As Long
    SocketType = m_ProtocolInfo.iSocketType
End Property


' *************************************************************************************************
' Minimum address length, in bytes.
' *************************************************************************************************
Public Property Get MinSockAddr() As Long
    MinSockAddr = m_ProtocolInfo.iMinSockAddr
End Property


' *************************************************************************************************
' Maximum address length, in bytes.
' *************************************************************************************************
Public Property Get MaxSockAddr() As Long
    MaxSockAddr = m_ProtocolInfo.iMaxSockAddr
End Property


' *************************************************************************************************
' Value to pass as the address family parameter to the socket/WSASocket function in order to open a
' socket for this protocol. This value also uniquely defines the structure of protocol addresses
' sockaddrs used by the protocol.
' *************************************************************************************************
Public Property Get AddressFamily() As AddressFamilyType
    AddressFamily = m_ProtocolInfo.iAddressFamily
End Property


' *************************************************************************************************
' Protocol version identifier.
' *************************************************************************************************
Public Property Get Version() As CWinsockVersion
    Set Version = New CWinsockVersion
    Version.IntegerVersion = m_ProtocolInfo.iVersion
End Property


' *************************************************************************************************
' Unique identifier assigned by the WS2_32.DLL for each WSAPROTOCOL_INFO structure.
' *************************************************************************************************
Public Property Get CatalogEntryId() As Long
    CatalogEntryId = m_ProtocolInfo.dwCatalogEntryId
End Property


' *************************************************************************************************
' Globally unique identifier assigned to the provider by the service provider vendor. This value is
' useful for instances where more than one service provider is able to implement a particular
' protocol. An application may use the dwProviderId value to distinguish between providers that
' might otherwise be indistinguishable.
' *************************************************************************************************
Public Property Get ProviderId() As String
    ProviderId = GuidToString(m_ProtocolInfo.ProviderId)
End Property


' *************************************************************************************************
' Indicates that a value of zero in the protocol parameter of socket or WSASocket matches this
' protocol entry.
' *************************************************************************************************
Public Property Get MatchesProtocolZero() As Boolean
    MatchesProtocolZero = m_ProtocolInfo.dwProviderFlags And ProtocolInfoProviderFlags.PFlags_MatchesProtocolZero
End Property


' *************************************************************************************************
' Indicates that this is the recommended or most frequently used entry for a protocol that is
' capable of implementing multiple behaviors.
' *************************************************************************************************
Public Property Get RecommendedProtoEntry() As Boolean
    RecommendedProtoEntry = m_ProtocolInfo.dwProviderFlags And ProtocolInfoProviderFlags.PFlags_RecommendedProtocolEntry
End Property


' *************************************************************************************************
' Indicates that this is one of two or more entries for a single protocol (from a given provider)
' which is capable of implementing multiple behaviors. An example of this is SPX which, on the
' receiving side, can behave either as a message-oriented or a stream-oriented protocol.
' *************************************************************************************************
Public Property Get MultipleProtoEntries() As Boolean
    MultipleProtoEntries = m_ProtocolInfo.dwProviderFlags And ProtocolInfoProviderFlags.PFlags_MultipleProtocolEntries
End Property


' *************************************************************************************************
' The MSG_PARTIAL flag is supported in WSASend and WSASendTo.
' *************************************************************************************************
Public Property Get PartialMessage() As Boolean
    PartialMessage = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_PartialMessage
End Property


' *************************************************************************************************
' Socket descriptors returned by the provider are operating system Installable File System (IFS)
' handles.
' *************************************************************************************************
Public Property Get UsesInstallableFileSystemHandles() As Boolean
    UsesInstallableFileSystemHandles = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_IFSHandles
End Property


' *************************************************************************************************
' Protocol is unidirectional in the recv direction.
' *************************************************************************************************
Public Property Get UnidirectionalRecv() As Boolean
    UnidirectionalRecv = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_UniRecv
End Property


' *************************************************************************************************
' Protocol is unidirectional in the send direction.
' *************************************************************************************************
Public Property Get UnidirectionalSend() As Boolean
    UnidirectionalSend = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_UniSend
End Property


' *************************************************************************************************
' Supports quality of service requests.
' *************************************************************************************************
Public Property Get QoSSupport() As Boolean
    QoSSupport = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_QOSSupported
End Property


' *************************************************************************************************
' Indicates whether the data plane is rooted (value = 1) or nonrooted (value = 0).
' *************************************************************************************************
Public Property Get MultipointDataPlane() As Boolean
    MultipointDataPlane = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_MultipointDataPlane
End Property


' *************************************************************************************************
' Indicates whether the control plane is rooted (value = 1) or nonrooted (value = 0).
' *************************************************************************************************
Public Property Get MultipointControlPlane() As Boolean
    MultipointControlPlane = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_MultipointControlPlane
End Property


' *************************************************************************************************
'  Supports a multipoint or multicast mechanism.
' *************************************************************************************************
Public Property Get SupportMultipoint() As Boolean
    SupportMultipoint = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_SupportMultipoint
End Property


' *************************************************************************************************
' Supports a broadcast mechanism.
' *************************************************************************************************
Public Property Get SupportBroadcast() As Boolean
    SupportBroadcast = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_SupportBroadcast
End Property


' *************************************************************************************************
' Supports disconnect data.
' *************************************************************************************************
Public Property Get DisconnectData() As Boolean
    DisconnectData = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_DisconnectData
End Property


' *************************************************************************************************
' Supports connect data.
' *************************************************************************************************
Public Property Get ConnectData() As Boolean
    ConnectData = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_ConnectData
End Property


' *************************************************************************************************
' Supports expedited (urgent) data.
' *************************************************************************************************
Public Property Get ExpeditedData() As Boolean
    ExpeditedData = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_ExpeditedData
End Property


' *************************************************************************************************
' Supports two-phase (graceful) close. If not set, only abortive closes are performed.
' *************************************************************************************************
Public Property Get GracefulClose() As Boolean
    GracefulClose = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_GracefulClose
End Property


' *************************************************************************************************
' A message-oriented protocol, but message boundaries are ignored for all receipts. This is
' convenient when an application does not desire message framing to be done by the protocol.
' *************************************************************************************************
Public Property Get PseudoStream() As Boolean
    PseudoStream = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_PseudoStream
End Property


' *************************************************************************************************
' Honors message boundaries—as opposed to a stream-oriented protocol where there is no concept of
' message boundaries.
' *************************************************************************************************
Public Property Get MessageOriented() As Boolean
    MessageOriented = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_MessageOriented
End Property


' *************************************************************************************************
' Guarantees that data only arrives in the order in which it was sent and that it is not
' duplicated. This characteristic does not necessarily mean that the data is always delivered,
' but that any data that is delivered is delivered in the order in which it was sent.
' *************************************************************************************************
Public Property Get GuaranteedOrder() As Boolean
    GuaranteedOrder = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_GuaranteedOrder
End Property


' *************************************************************************************************
' Guarantees that all data sent will reach the intended destination.
' *************************************************************************************************
Public Property Get GuaranteedDelivery() As Boolean
    GuaranteedDelivery = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_GuaranteedDelivery
End Property


' *************************************************************************************************
' Provides connectionless (datagram) service. If not set, the protocol supports connection-oriented
' data transfer.
' *************************************************************************************************
Public Property Get Connectionless() As Boolean
    Connectionless = m_ProtocolInfo.dwServiceFlags1 And ProtocolServiceFlags1.ServFlag1_Connectionless
End Property


' *************************************************************************************************
' The protocol chain.
' *************************************************************************************************
Public Property Get ProtocolChain() As CProtocolChain
    
    ' Create a new protocol chain class
    Set ProtocolChain = New CProtocolChain

    ' Copy the protocol chain structure into the class
    Call ProtocolChain.SetProtocolChainInfo(VarPtr(m_ProtocolInfo.ProtocolChain))

End Property
