VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProtocolEnumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CProtocolInfo"
Attribute VB_Ext_KEY = "Member0" ,"CProtocolInfo"
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


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)


' *************************************************************************************************
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Exception As CWinsockException)


'local variable to hold collection
Private m_colProtocols As Collection


' *************************************************************************************************
' Get an individual protocol item.
' *************************************************************************************************
Public Property Get Item(vntIndexKey As Variant) As CProtocolInfo
  On Error Resume Next 'Incase they pass an invalid index
    Set Item = m_colProtocols(vntIndexKey)
End Property


' *************************************************************************************************
' Get the protocol count.
' *************************************************************************************************
Public Property Get Count() As Long
    Count = m_colProtocols.Count
End Property


Public Sub Initialize()

  Dim lngRetVal       As Long
  Dim lngBufferSize   As Long
  Dim Buffer()        As Byte
  Dim ProtocolInfo    As API_WSAPROTOCOL_INFO
  Dim NewMember       As CProtocolInfo
  Dim ErrorObject     As CWinsockException
  Dim i               As Integer
    
    Set m_colProtocols = New Collection

    ReDim Buffer(1)

    ' First call the WSAEnumProtocols function to get the buffer size
    lngRetVal = api_WSAEnumProtocols(0&, Buffer(0), lngBufferSize)

    'If lngRetVal = SOCKET_ERROR Then
    '    Set ErrorObject = New CWinsockException
    '    ErrorObject.Source = "CProtocolEnumeration.Initialize"
    '    RaiseEvent OnError(ErrorObject)
    '    Exit Sub
    'End If

    ' Resize the buffer to the correct size
    ReDim Buffer(lngBufferSize - 1)

    ' Call the WSAEnumProtocols function again to get the actual protocols
    lngRetVal = api_WSAEnumProtocols(0&, Buffer(0), lngBufferSize)
    
    If lngRetVal = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CProtocolEnumeration.Class_Initialize"
        RaiseEvent OnError(ErrorObject)
        Exit Sub
    End If
    
    ' Loop through each protocol
    For i = 0 To lngRetVal - 1
    
        ' Copy the protocol info
        CopyMemory ProtocolInfo, Buffer(i * Len(ProtocolInfo)), Len(ProtocolInfo)
        
        ' Create a new protocol info class
        Set NewMember = New CProtocolInfo
        
        ' Initialize the class by passing the current protocol info structure to it
        NewMember.GetProtocolByProtocolInfoType VarPtr(ProtocolInfo)
        
        ' Add the new protocol to the collection
        m_colProtocols.Add NewMember
        
        ' Dipose of this now
        Set NewMember = Nothing
    Next i
End Sub


Private Sub Class_Terminate()
    Set m_colProtocols = Nothing
End Sub

