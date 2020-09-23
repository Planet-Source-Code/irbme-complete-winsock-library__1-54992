VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CServiceEntry"
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


' A private copy of the service entry structure from which properties can read the data from
Private m_ServiceEntry As API_SERVENT


' *************************************************************************************************
' Description: Set the service entry info. lpServiceEntry MUST be a valid pointer to a
'              API_SERVENT structure or else a general protection fault is likely to occur
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    09/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub GetServiceByServent(ByVal lpServiceEntry As Long)
    If lpServiceEntry <= 0 Then Exit Sub
    RtlMoveMemory ByVal VarPtr(m_ServiceEntry), lpServiceEntry, LenB(m_ServiceEntry)
End Sub


' *************************************************************************************************
' The service name. Setting the service name will also attempt to get the service info by that
' name.
' *************************************************************************************************

Public Property Get ServiceName() As String
    ServiceName = StringFromPointer(m_ServiceEntry.s_name)
End Property

Public Property Let ServiceName(Value As String)
    
  Dim lpServiceEntry As Long
  Dim ErrorObject    As CWinsockException
    
    m_ServiceEntry.s_aliases = 0
    m_ServiceEntry.s_name = 0
    m_ServiceEntry.s_port = 0
    m_ServiceEntry.s_proto = 0
    
    If Value = vbNullString Then Exit Property

    
    lpServiceEntry = api_GetServByName(Value, 0&)
    
    If lpServiceEntry = 0 Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CServiceEntry.Let ServiceName"
        RaiseEvent OnError(ErrorObject)
        Exit Property
    Else
        GetServiceByServent lpServiceEntry
    End If
    
End Property


' *************************************************************************************************
' The protocol name such as "tcp".
' *************************************************************************************************

Public Property Get ProtocolName() As String
    ProtocolName = StringFromPointer(m_ServiceEntry.s_proto)
End Property


' *************************************************************************************************
' A collection of strings. The aliases that the service is also known as.
' *************************************************************************************************
Public Property Get Aliases() As Collection
    Set Aliases = PtrArrayToStrCollection(m_ServiceEntry.s_aliases)
End Property


' *************************************************************************************************
' The port. Setting the port will also attempt to get the service info by tht port.
' *************************************************************************************************

Public Property Get Port() As Long
    Port = IntegerToUnsigned(api_ntohs(m_ServiceEntry.s_port))
End Property

Public Property Let Port(Value As Long)
    
  Dim lpServiceEntry As Long
  Dim ErrorObject    As CWinsockException
    
    
    m_ServiceEntry.s_aliases = 0
    m_ServiceEntry.s_name = 0
    m_ServiceEntry.s_port = 0
    m_ServiceEntry.s_proto = 0
    
    If Value = 0 Then Exit Property
    
    lpServiceEntry = api_GetServByPort(api_htons(UnsignedToInteger(Value)), 0&)
    
    If lpServiceEntry = 0 Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CServiceEntry.Let Port"
        RaiseEvent OnError(ErrorObject)
        Exit Property
    Else
        GetServiceByServent lpServiceEntry
    End If
    
End Property
