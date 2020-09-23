VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterfaceEnumeration"
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


' A collection of protocols
Private m_colInterfaces As Collection


' *************************************************************************************************
' Get an individual protocol item.
' *************************************************************************************************
Public Property Get Item(Key As Variant) As CInterface
  On Error Resume Next ' Just incase index doesn't exist
    Set Item = m_colInterfaces(Key)
End Property


' *************************************************************************************************
' Get the protocol count.
' *************************************************************************************************
Public Property Get Count() As Long
    Count = m_colInterfaces.Count
End Property


Private Sub Class_Initialize()

  Dim InBuffer          As Long
  Dim OutBuffer(0 To 9) As API_INTERFACE_INFO
  Dim BytesReturned     As Long
  Dim NumberInterfaces  As Long, i As Long
  Dim lngSocketHandle   As Long
  Dim Interface         As CInterface
  Dim ErrorObject       As CWinsockException

    ' Attempt to create the socket
    lngSocketHandle = api_Socket(AF_INET, SOCK_RAW, IPPROTO_RAW)

    ' Check for an error
    If lngSocketHandle = INVALID_SOCKET Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CInterfaceEnumeration.Class_Initialize"
        RaiseEvent OnError(ErrorObject)
        Exit Sub
    End If

    ' Call WSAIoctl in order to get an interface list.
    ' The outbuffer is an array of interfaces and BytesReturned contains
    ' the number of bytes of the outbuffer which was used
    If api_WSAIoctl(lngSocketHandle, SIO_GET_INTERFACE_LIST, InBuffer, LenB(InBuffer), OutBuffer(0), LenB(OutBuffer(0)) * 10, BytesReturned, 0&, 0&) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CInterfaceEnumeration.Class_Initialize"
        RaiseEvent OnError(ErrorObject)
        Exit Sub
    End If

    ' Calculate the number of interfaces
    NumberInterfaces = BytesReturned / LenB(OutBuffer(0))
    
    Set m_colInterfaces = New Collection

    ' Loop through each interface
    For i = 0 To NumberInterfaces - 1
        Set Interface = New CInterface

        Interface.GetInterfaceByStructPtr VarPtr(OutBuffer(i))
        m_colInterfaces.Add Interface

        Set Interface = Nothing
    Next i

End Sub


Private Sub Class_Terminate()
    Set m_colInterfaces = Nothing
End Sub
