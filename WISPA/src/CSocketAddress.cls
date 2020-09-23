VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocketAddress"
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


Private m_ip4Addr As API_SOCKADDR_IN
Private m_cIPAddress As CIP4Address


' *************************************************************************************************
' Description: Set the sockaddr info. lpSockAddr MUST be a valid pointer to a
'              API_SOCKADDR_IN structure or else a general protection fault is likely to occur
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    05/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub GetBySockAddr(ByVal lpSockaddr As Long)
    If lpSockaddr <= 0 Then Exit Sub

    Set m_cIPAddress = New CIP4Address
    RtlMoveMemory ByVal VarPtr(m_ip4Addr), lpSockaddr, LenB(m_ip4Addr)

    m_cIPAddress.Address = m_ip4Addr.sin_addr.S_addr

End Sub


' *************************************************************************************************
' The IP Version 4 address.
' *************************************************************************************************
Public Property Get IPAddress() As CIP4Address
    Set IPAddress = m_cIPAddress
End Property

Public Property Set IPAddress(Value As CIP4Address)
    Set m_cIPAddress = Value
    m_ip4Addr.sin_addr.S_addr = Value.Address
End Property
    


' *************************************************************************************************
' The port.
' *************************************************************************************************
Public Property Get Port() As Long
    Port = IntegerToUnsigned(m_ip4Addr.sin_port)
End Property

Public Property Let Port(Value As Long)
    m_ip4Addr.sin_port = UnsignedToInteger(Value)
End Property


' *************************************************************************************************
' The address family.
' *************************************************************************************************
Public Property Get AddressFamily() As AddressFamilyType
    AddressFamily = m_ip4Addr.sin_family
End Property

Public Property Let AddressFamily(Value As AddressFamilyType)
    m_ip4Addr.sin_family = Value
End Property


Private Sub Class_Terminate()
    Set m_cIPAddress = Nothing
End Sub
