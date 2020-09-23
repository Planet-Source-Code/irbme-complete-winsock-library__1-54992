VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProtocolEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


' A private copy of the protocol entry structure from which properties can read the data from
Private m_ProtocolEntry As API_PROTOENT


' *************************************************************************************************
' Description: Set the protocol entry info. lpProtocolEntry MUST be a valid pointer to a
'              API_PROTOENT structure or else a general protection fault is likely to occur
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    09/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub GetProtocolByProtoent(ByVal lpProtocolEntry As Long)
    If lpProtocolEntry <= 0 Then Exit Sub
    RtlMoveMemory ByVal VarPtr(m_ProtocolEntry), lpProtocolEntry, LenB(m_ProtocolEntry)
End Sub


' *************************************************************************************************
' The service name. Setting the service name will also attempt to get the service info by that
' name.
' *************************************************************************************************

Public Property Get ProtocolName() As String
    ProtocolName = StringFromPointer(m_ProtocolEntry.p_name)
End Property

Public Property Let ProtocolName(Value As String)
    
  Dim lpProtocolEntry As Long
  Dim ErrorObject     As CWinsockException
    
    m_ProtocolEntry.p_aliases = 0
    m_ProtocolEntry.p_name = 0
    m_ProtocolEntry.p_proto = 0
    
    If Value = vbNullString Then Exit Property
    
    lpProtocolEntry = api_GetProtoByName(Value)
    
    If lpProtocolEntry = 0 Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CProtocolEntry.Let ProtocolName"
        RaiseEvent OnError(ErrorObject)
        Exit Property
    Else
        GetProtocolByProtoent lpProtocolEntry
    End If
    
End Property



' *************************************************************************************************
' A collection of strings. The aliases that the service is also known as.
' *************************************************************************************************
Public Property Get Aliases() As Collection
    Set Aliases = PtrArrayToStrCollection(m_ProtocolEntry.p_aliases)
End Property


' *************************************************************************************************
' The port. Setting the port will also attempt to get the service info by tht port.
' *************************************************************************************************

Public Property Get ProtocolNumber() As Long
    ProtocolNumber = m_ProtocolEntry.p_proto
End Property

Public Property Let ProtocolNumber(Value As Long)
    
  Dim lpProtocolEntry As Long
  Dim ErrorObject     As CWinsockException
    
    m_ProtocolEntry.p_aliases = 0
    m_ProtocolEntry.p_name = 0
    m_ProtocolEntry.p_proto = 0

     If Value = 0 Then Exit Property

    lpProtocolEntry = api_GetProtoByNumber(Value)
    
    If lpProtocolEntry = 0 Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CProtocolEntry.Let ProtocolNumber"
        RaiseEvent OnError(ErrorObject)
        Exit Property
    Else
        GetProtocolByProtoent lpProtocolEntry
    End If
    
End Property
