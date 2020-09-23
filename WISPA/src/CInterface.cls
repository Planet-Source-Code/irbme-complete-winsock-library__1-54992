VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterface"
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


' The local copy
Private m_Interface As API_INTERFACE_INFO


' *************************************************************************************************
' Description: Get the interface by a pointer to an API_INTERFACE_INFO structure.
'              lpInterfaceInfo MUST be a valid pointer to an API_INTERFACE_INFO structure or else a
'              general protection fault is likely to occur
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    05/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub GetInterfaceByStructPtr(ByVal lpInterfaceInfo As Long)
    If lpInterfaceInfo <= 0 Then Exit Sub
    RtlMoveMemory ByVal VarPtr(m_Interface), lpInterfaceInfo, LenB(m_Interface)
End Sub


' *************************************************************************************************
' The address of the interface.
' *************************************************************************************************
Public Property Get Address() As CSocketAddress
    Set Address = New CSocketAddress
    Address.GetBySockAddr VarPtr(m_Interface.iiAddress.AddressIn)
End Property


' *************************************************************************************************
' The address used to broadcast on the interface. Only valid if broadcasting is supported.
' *************************************************************************************************
Public Property Get BroadcastAddress() As CSocketAddress
    Set BroadcastAddress = New CSocketAddress
    BroadcastAddress.GetBySockAddr VarPtr(m_Interface.iiBroadcastAddress.AddressIn)
End Property


' *************************************************************************************************
' The subnet mask address of the interface.
' *************************************************************************************************
Public Property Get SubnetMask() As CSocketAddress
    Set SubnetMask = New CSocketAddress
    SubnetMask.GetBySockAddr VarPtr(m_Interface.iiNetmask.AddressIn)
End Property


' *************************************************************************************************
' Whether or not the interface is up (enabled).
' *************************************************************************************************
Public Property Get InterfaceUp() As Boolean
    InterfaceUp = m_Interface.iiFlags And IFF_UP
End Property


' *************************************************************************************************
' Whether broadcasting is supported.
' *************************************************************************************************
Public Property Get BroadcastSupported() As Boolean
    BroadcastSupported = m_Interface.iiFlags And IFF_BROADCAST
End Property


' *************************************************************************************************
' Whether or not multicasting is supported
' *************************************************************************************************
Public Property Get MulticastSupported() As Boolean
    MulticastSupported = m_Interface.iiFlags And IFF_MULTICAST
End Property


' *************************************************************************************************
' Is this the loopback interface.
' *************************************************************************************************
Public Property Get IsLoopbackInterface() As Boolean
    IsLoopbackInterface = m_Interface.iiFlags And IFF_LOOPBACK
End Property


' *************************************************************************************************
' Is this a point to point link interface?
' *************************************************************************************************
Public Property Get IsPointToPointLink() As Boolean
    IsPointToPointLink = m_Interface.iiFlags And IFF_POINTTOPOINT
End Property
