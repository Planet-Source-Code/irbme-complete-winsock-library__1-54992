VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinsockStartup"
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


' The WSAData structure used to store the information returned from WSAStrartup
Private m_WinsockData As API_WSADATA

' If the wnisock API was successfuly started then this will be true
Private m_bSuccessfulStartup As Boolean


Private m_Version As CWinsockVersion
Private m_HighVersion As CWinsockVersion


' *************************************************************************************************
' Version of the Windows Sockets specification that the Ws2_32.dll expects the caller to use.
' *************************************************************************************************
Public Property Get Version() As CWinsockVersion
    Set Version = m_Version
End Property


' *************************************************************************************************
' Highest version of the Windows Sockets specification that this .dll can support
' (also encoded as above). Normally this is the same as wVersion.
' *************************************************************************************************
Public Property Get HighestVersion() As CWinsockVersion
    Set HighestVersion = m_HighVersion
End Property


' *************************************************************************************************
' Null-terminated ASCII string into which the Ws2_32.dll copies a description of the Windows
' Sockets implementation. The text (up to 256 characters in length) can contain any
' characters except control and formatting characters: the most likely use that an application
' can put this to is to display it (possibly truncated) in a status message.
' *************************************************************************************************
Public Property Get Description() As String
    Description = Trim$(Replace$(m_WinsockData.szDescription, Chr$(0), " "))
End Property


' *************************************************************************************************
' Null-terminated ASCII string into which the WSs2_32.dll copies relevant status or
' configuration information. The Ws2_32.dll should use this parameter only if the information
' might be useful to the user or support staff: it should not be considered as an extension
' of the szDescription parameter.
' *************************************************************************************************
Public Property Get SystemStatus() As String
    SystemStatus = Trim$(Replace$(m_WinsockData.szSystemStatus, Chr$(0), " "))
    
    ' Stolen from the winsock implementation in the mono project - thought it was a nice touch :)
    If SystemStatus = "Running" Then
        SystemStatus = "Groovy"
    End If
    
End Property


' *************************************************************************************************
' Retained for backward compatibility, but should be ignored for Windows Sockets version 2
' and later, as no single value can be appropriate for all underlying service providers.
' *************************************************************************************************
Public Property Get MaxSockets() As Integer
    MaxSockets = m_WinsockData.iMaxSockets
End Property


' *************************************************************************************************
' Has the winsock API been successfuly started up
' *************************************************************************************************
Public Property Get SuccessfulyStartedUp() As Boolean
    SuccessfulyStartedUp = m_bSuccessfulStartup
End Property


Private Sub Class_Initialize()

  Dim ErrorObject   As CWinsockException
  Dim ErrorCode     As ErrorCodeType


    ' No point starting another version of the winsock dll if it's already been started
    ' once already
    If WSAStartupCount = 1 Then Exit Sub

    ' Attempt to start up the winsock API (version 2.2)
    ErrorCode = modWinsock.api_WSAStartup(modWinsockUtils.MakeVersion(2, 2), m_WinsockData)
        
    ' If there was an error then display it
    If ErrorCode <> ErrorCodeType.NoError Then

        Set ErrorObject = New CWinsockException
        ErrorObject.ErrorCode = ErrorCode
        ErrorObject.Source = "CWinsockStartup.Class_Initialize"
            
        RaiseEvent OnError(ErrorObject)
    Else
    
        ' Set up the version properties
        Set m_Version = New CWinsockVersion
        Set m_HighVersion = New CWinsockVersion
        
        m_Version.IntegerVersion = m_WinsockData.wVersion
        m_HighVersion.IntegerVersion = m_WinsockData.wHighVersion
        
        ' And some status variables
        m_bSuccessfulStartup = True
        WSAStartupCount = WSAStartupCount + 1
    End If

End Sub


Private Sub Class_Terminate()
    If m_bSuccessfulStartup Then
        modWinsock.api_WSACleanup
        WSAStartupCount = WSAStartupCount - 1
    End If
End Sub

