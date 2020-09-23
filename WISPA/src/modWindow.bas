Attribute VB_Name = "modWindow"
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
' Subclassing Win32 API calls.
' *************************************************************************************************
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' *************************************************************************************************
' Used to associate and retrieve custom data with a window handle.
' *************************************************************************************************
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

' *************************************************************************************************
' Window creation and destruction.
' *************************************************************************************************
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

' *************************************************************************************************
' Copy memory.
' *************************************************************************************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Private Const GWL_WNDPROC = (-4)


' *************************************************************************************************
' Description:  Creates a new window solely for the use of subclassing. Pass a pointer to
'               a CWindow class instance into the lpClass parameter. This can be gotten by using
'               ObjPtr(ClassInstance) or ObjPtr(Me) from inside the class. The class pointer
'               is associated with the handle for internal use. A handle to the window is returned.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function CreateSocketWindow(lpClass As Long) As Long

  Dim hwnd As Long
  
    If lpClass <= 0 Then
        Err.Raise -1, "modWindow.CreateSocketWindow", "The class pointer must be a valid pointer to a CWindow class!"
        Exit Function
    End If
    
    ' Create the most basic type of window
    hwnd = CreateWindowEx(0&, "STATIC", 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, ByVal 0&)

    ' If window creation was successful then
    If hwnd > 0 Then
    
        ' Associate the class pointer with the window handle
        If SetProp(hwnd, "Int32Ptr", lpClass) = 0 Then
            Err.Raise -1, "modWinsock.CreateSocketWindow", "Could not associate the class pointer with the window handle. Last error: " & Err.LastDllError
            Exit Function
        End If
        
        ' Subclass the window
        If SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc) = 0 Then
            Err.Raise -1, "modWinsock.CreateSocketWindow", "Could not subclass the window. Last error: " & Err.LastDllError
            Exit Function
        End If
    Else
        Err.Raise -1, "modWinsock.CreateSocketWindow", "Failed to create the window. Last error: " & Err.LastDllError
    End If
    
    ' Return the handle
    CreateSocketWindow = hwnd
    
End Function



' *************************************************************************************************
' Description:  The window proc. All messages for all windows are directed here. The class pointer
'               associated with the window handle is extracted and used to obtain the instance of
'               the class. The classes WindowProc method is then called, and the relevant data
'               passed to it.
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim objIllegal   As Object
  Dim objWindow    As CWindow
  Dim lpClassPtr   As Long
  Dim bHandled     As Boolean

    ' Extract the class pointer
    lpClassPtr = CLng(GetProp(hwnd, "Int32Ptr"))

    If lpClassPtr = 0 Then
    
        ' This is probably a bad idea in a windowproc but it's very very unlikely to happen
        ' unless the client is doing some funny stuff.
        Err.Raise -1, "modWinsock.WindowProc", "Failed to get the class pointer from the window handle. Last error: " & Err.LastDllError
    End If

    ' Obtain an illegal, uncounted reference to the class
    CopyMemory objIllegal, lpClassPtr, 4&
    
    ' Make it legal
    Set objWindow = objIllegal

    ' Call the window proc method of the class
    WindowProc = objWindow.WindowProc(hwnd, uMsg, wParam, lParam, bHandled)
    
    ' Get rid of the illegal reference
    CopyMemory objIllegal, 0&, 4&
    Set objIllegal = Nothing

    ' If this message isn't handled then handle it with default message processing
    If Not bHandled Then WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    
End Function
