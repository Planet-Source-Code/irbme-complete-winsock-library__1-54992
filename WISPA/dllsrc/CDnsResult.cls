VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDnsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

' Fully qualified domain name
Private m_strFullyQualifiedDomainName   As String

' array of aliases
Private m_Aliases                       As Collection


' Collection of addresses
Private m_AddressList                   As CIP4Collection


' *************************************************************************************************
' The fully qualified domain name.
' *************************************************************************************************
Public Property Get FullyQualifiedDomainName() As String
    FullyQualifiedDomainName = m_strFullyQualifiedDomainName
End Property

Public Property Let FullyQualifiedDomainName(Value As String)
    m_strFullyQualifiedDomainName = Value
End Property


' *************************************************************************************************
' Alias list
' *************************************************************************************************
Public Property Get Aliases() As Collection
    Set Aliases = m_Aliases
End Property

Public Property Set Aliases(Value As Collection)
    Set m_Aliases = Value
End Property


' *************************************************************************************************
' IP Address list
' *************************************************************************************************
Public Property Get AddressList() As CIP4Collection
    Set AddressList = m_AddressList
End Property

Public Property Set AddressList(Value As CIP4Collection)
    Set m_AddressList = Value
End Property
