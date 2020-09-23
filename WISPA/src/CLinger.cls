VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLinger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

' Whether or not socket linger is enabled
Private m_LingerEnabled As Boolean

' If it is enabled, then the timeout value
Private m_LingerTimeout As Long



' *************************************************************************************************
' Description: Set the linger info. lpLinger MUST be a valid pointer to a
'              API_LINGER structure or else a general protection fault is likely to occur
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    04/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub SetLinger(ByVal lpLinger As Long)
  Dim Linger As API_LINGER

    If lpLinger <= 0 Then Exit Sub

    RtlMoveMemory ByVal VarPtr(Linger), lpLinger, LenB(Linger)
    
    m_LingerEnabled = Linger.l_onoff <> 0
    m_LingerTimeout = Linger.l_linger

End Sub


' *************************************************************************************************
' Linger is enabled.
' *************************************************************************************************
Public Property Get LingerEnabled() As Boolean
    LingerEnabled = m_LingerEnabled
End Property

Public Property Let LingerEnabled(Value As Boolean)
    m_LingerEnabled = Value
End Property


' *************************************************************************************************
' Timeout.
' *************************************************************************************************
Public Property Get TimeOut() As Long
    TimeOut = m_LingerTimeout
End Property

Public Property Let TimeOut(Value As Long)
    m_LingerTimeout = Value
End Property
