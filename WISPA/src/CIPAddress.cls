VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIP4Address"
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


Private m_lngAddr As Long
Private m_strAddr As String


' *************************************************************************************************
' Some commonly used addresses
' *************************************************************************************************

Public Property Get AnyAddress() As CIP4Address
    Set AnyAddress = New CIP4Address
    AnyAddress.Address = INADDR_ANY
End Property

Public Property Get LoopbackAddress() As CIP4Address
    Set AnyAddress = New CIP4Address
    AnyAddress.Address = INADDR_LOOPBACK
End Property

Public Property Get BroadcastAddress() As CIP4Address
    Set AnyAddress = New CIP4Address
    AnyAddress.Address = INADDR_BROADCAST
End Property

Public Property Get NoAddress() As CIP4Address
    Set AnyAddress = New CIP4Address
    AnyAddress.Address = INADDR_NONE
End Property


' *************************************************************************************************
' The long representation of the address
' *************************************************************************************************
Public Property Get Address() As Long
    Address = m_lngAddr
End Property

Public Property Let Address(Value As Long)
    m_lngAddr = Value
    m_strAddr = GetStrIPFromLong(Value)
End Property


' *************************************************************************************************
' A more human readable representation of the address. Returns the address in standard dotted IP
' format.
' *************************************************************************************************
Public Property Get StringAddress() As String
    StringAddress = m_strAddr
End Property

Public Property Let StringAddress(Value As String)
    m_strAddr = Value
    m_lngAddr = api_ntohl(GetLngIPFromStr(Value))
End Property
