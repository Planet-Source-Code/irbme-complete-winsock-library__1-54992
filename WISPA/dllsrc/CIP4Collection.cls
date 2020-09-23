VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIP4Collection"
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


' *************************************************************************************************
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Exception As CWinsockException)


' A collection of protocols
Private m_colIPs As Collection


' *************************************************************************************************
' Add a new item
' *************************************************************************************************
Public Sub Add(NewItem As CIP4Address)
    m_colIPs.Add NewItem
End Sub


' *************************************************************************************************
' Get an individual protocol item.
' *************************************************************************************************
Public Property Get Item(Index As Variant) As CIP4Address
    Set Item = m_colIPs(Index)
End Property


' *************************************************************************************************
' Get the protocol count.
' *************************************************************************************************
Public Property Get Count() As Long
    Count = m_colIPs.Count
End Property


Private Sub Class_Initialize()
    Set m_colIPs = New Collection
End Sub


Private Sub Class_Terminate()
    Set m_colIPs = Nothing
End Sub

