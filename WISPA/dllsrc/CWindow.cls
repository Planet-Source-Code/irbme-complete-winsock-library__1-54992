VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindow"
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
' Used to create a new unique message number which can be posted to wndprocs without interfering
' with existing ones.
' *************************************************************************************************
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long


' *************************************************************************************************
' The MessageArrived event is raised whenever a new message arrives at the wndproc.
' lngRetVal is the value to be returned to the OS by the WndProc.
' If bHandled is set to false then the message is processed by the default message handler
' otherwise it's not processed. Note that if bHandled = false then RetVal is meaningless.
' *************************************************************************************************
Public Event MessageArrived(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lngRetVal As Long, bHandled As Boolean)

' The window handle
Private m_hWnd As Long

' The winsock message
Private m_WinsockMessage As Long


' *************************************************************************************************
' Winsock message
' *************************************************************************************************
Public Property Get WinsockMessage() As Long
    WinsockMessage = m_WinsockMessage
End Property


' *************************************************************************************************
' Window Handle
' *************************************************************************************************
Public Property Get WindowHandle() As Long
    WindowHandle = m_hWnd
End Property


' *************************************************************************************************
' This method is called by the main window procedure from modWindow whenever a message arrives.
' *************************************************************************************************
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef bHandled As Boolean) As Long

  Dim RetVal As Long
  
    If uMsg <> m_WinsockMessage Then
        bHandled = False
        Exit Function
    End If
  
    ' Raise the event
    RaiseEvent MessageArrived(hwnd, uMsg, wParam, lParam, RetVal, bHandled)
    WindowProc = RetVal

End Function


Private Sub Class_Initialize()
    
    ' Create the socket message
    m_WinsockMessage = RegisterWindowMessage(App.EXEName & ":" & App.hInstance & ":Winsock:" & CStr(ObjPtr(Me)))
    
    ' Create a new socket window. Pass the class pointer so it knows who we are and can call
    ' out WindowProc method successfuly
    m_hWnd = CreateSocketWindow(ObjPtr(Me))
End Sub


Private Sub Class_Terminate()
    ' Very important. Without this, the application goes kaboom very quickly.
    If m_hWnd > 0 Then DestroyWindow m_hWnd
End Sub
