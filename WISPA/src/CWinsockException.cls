VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinsockException"
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

' The source of the error
Private m_strSource As String

' The error code
Private m_errErrorCode As ErrorCodeType


' *************************************************************************************************
' The last winsock error which occured.
' *************************************************************************************************
Public Property Get LastWinsockError() As ErrorCodeType
    LastWinsockError = Err.LastDllError
End Property

Public Property Let LastWinsockError(Value As ErrorCodeType)
    modWinsock.api_WSASetLastError Value
End Property


' *************************************************************************************************
' The source of the error.
' *************************************************************************************************
Public Property Get Source() As String
    Source = m_strSource
End Property

Public Property Let Source(Value As String)
    m_strSource = Value
End Property


' *************************************************************************************************
' The actual error code.
' *************************************************************************************************
Public Property Get ErrorCode() As ErrorCodeType
    ErrorCode = m_errErrorCode
End Property

Public Property Let ErrorCode(Value As ErrorCodeType)
    m_errErrorCode = Value
End Property


' *************************************************************************************************
' A short description of the error
' *************************************************************************************************
Public Property Get ErrorDescription() As String
    ErrorDescription = GetErrorDescription(m_errErrorCode, DescriptionShort)
End Property


' *************************************************************************************************
' Description:  Display the error code.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Sub Display()
    
    MsgBox "The program encountered an error in " & m_strSource & vbCrLf & vbCrLf & _
           "Error Code: " & m_errErrorCode & _
           " (" & GetErrorDescription(m_errErrorCode, DescriptionWin32API) & ")" & vbCrLf & _
           "Description: " & GetErrorDescription(m_errErrorCode, DescriptionShort) & vbCrLf & vbCrLf & _
           GetErrorDescription(m_errErrorCode, DescriptionLong), vbOKOnly Or vbCritical, "Winsock Error"

End Sub


Private Sub Class_Initialize()

    ' Initialize the error code to the last dll error which occured. This saves most of
    ' clients from having to set the error code since in most cases the error is stored
    ' in the last dll error
    m_errErrorCode = Err.LastDllError
End Sub
