VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProtocolChain"
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
' Type of protocol chain
' *************************************************************************************************
Public Enum ProtocolChainType
    ' If the length of the chain is 0, this WSAPROTOCOL_INFO entry represents a layered protocol
    ' which has Windows Sockets 2 SPI as both its top and bottom edges.
    LayeredProtocol = LAYERED_PROTOCOL
    
    ' If the length of the chain equals 1, this entry represents a base protocol whose Catalog
    ' Entry identifier is in the dwCatalogEntryId member of the WSAPROTOCOL_INFO structure.
    BaseProtocol = BASE_PROTOCOL
    
    ' If the length of the chain is larger than 1, this entry represents a protocol chain which
    ' consists of one or more layered protocols on top of a base protocol.
    LayeredProtocolChain = &H2
End Enum


' The protocol entries
Private m_lngEntries() As Long

' The protocol type
Private m_lngType As ProtocolChainType


' *************************************************************************************************
' Description: Set the protocol chain info. lpProtocolChain MUST be a valid pointer to a
'              API_WSAPROTOCOLCHAIN structure or else a general protection fault is likely to occur
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    04/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub SetProtocolChainInfo(ByVal lpProtocolChain As Long)
  Dim ProtocolChain As API_WSAPROTOCOLCHAIN

    RtlMoveMemory ByVal VarPtr(ProtocolChain), lpProtocolChain, LenB(ProtocolChain)
    ReDim Preserve m_lngEntries(-1 To ProtocolChain.ChainLen - 1)

    Select Case ProtocolChain.ChainLen
        Case 0
           m_lngType = LayeredProtocol
        Case 1
            m_lngType = BaseProtocol
        Case Else
            m_lngType = LayeredProtocolChain
    End Select

End Sub


' *************************************************************************************************
' The protocol chain type.
' *************************************************************************************************
Public Property Get ChainType() As ProtocolChainType
    ChainType = m_lngType
End Property


' *************************************************************************************************
' The number of entries in the chain.
' *************************************************************************************************
Public Property Get ChainEntryCount() As Long
    ChainEntryCount = UBound(m_lngEntries) + 1
End Property


' *************************************************************************************************
' Get a single entry.
' *************************************************************************************************
Public Property Get ChainEntry(Index As Long) As Long
  On Error Resume Next 'Incase they pass an invalid index
    ChainEntry = m_lngEntries(Index)
End Property


Private Sub Class_Initialize()
    ReDim m_lngEntries(-1 To -1)
End Sub

Private Sub Class_Terminate()
    Erase m_lngEntries
End Sub
