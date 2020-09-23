VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinsockVersion"
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


' *************************************************************************************************
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Exception As CWinsockException)


Private m_intVersion As Integer


' *************************************************************************************************
' The major part of the version
' *************************************************************************************************
Public Property Get Major() As Byte
    Major = HiByte(m_intVersion)
End Property

Public Property Let Major(ByVal Value As Byte)
    m_intVersion = MakeWord(Value, Minor)
End Property


' *************************************************************************************************
' The minor part of the version
' *************************************************************************************************
Public Property Get Minor() As Byte
    Minor = LoByte(m_intVersion)
End Property

Public Property Let Minor(ByVal Value As Byte)
    m_intVersion = MakeWord(Major, Value)
End Property


' *************************************************************************************************
' The entire word representing the version
' *************************************************************************************************
Public Property Get IntegerVersion() As Integer
    IntegerVersion = m_intVersion
End Property

Public Property Let IntegerVersion(ByVal Value As Integer)
    m_intVersion = Value
End Property


' *************************************************************************************************
' The string representation of the version
' *************************************************************************************************
Public Property Get StringVersion() As String
    StringVersion = Major & "." & Minor
End Property

Public Property Let StringVersion(ByVal Value As String)

  Dim VersionParts() As String

    ' Makes sure it's has at least 2 character seperated by a dot
    If Value Like "*?.?*" Then
    
        ' Split up the 2 characters
        VersionParts = Split(Value, ".")
        
        If UBound(VersionParts) = 1 Then
        
            ' Make sure they're numeric and between 0 and 256
            If IsNumeric(VersionParts(0)) And IsNumeric(VersionParts(1)) Then
                If VersionParts(0) > 0 And VersionParts(0) < 256 And VersionParts(1) > 0 And VersionParts(1) < 256 Then
                    ' The version passed is fine, save it
                    m_intVersion = MakeWord(CByte(VersionParts(0)), CByte(VersionParts(1)))
                    Exit Property
                End If
            End If
        End If
    End If
    
    Err.Raise 0, "CWinsockVersion.Property Let StringVersion", "The version, " & Value & " is invalid"
    
End Property
