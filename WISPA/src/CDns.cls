VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


' Used instead of the default public RTLMoveMemory declaration, which uses pointers explicitly,
' in some places
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' Used to create a new unique message number which can be posted to wndprocs without interfering
' with existing ones.
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long


' *************************************************************************************************
' Raised whenever an error occurs.
' *************************************************************************************************
Public Event OnError(Exception As CWinsockException)

' *************************************************************************************************
' Raised whenever an async resolve is finished successfuly.
' *************************************************************************************************
Public Event HostResolved(HostEntry As CDnsResult)


' Used to create a new
Private WithEvents EventWindow As CWindow
Attribute EventWindow.VB_VarHelpID = -1

' The unique message id used in the window proc
Private ResolveHostMessage As Long

' The buffer to hold the hostent structure and all the other relevant data such as the
' address list, alias list, and host name
Private ResolveBuffer(MAXGETHOSTSTRUCT - 1) As Byte

' The task handle returned by the WSAGetXByY methods
Private TaskHandle As Long


' *************************************************************************************************
' Description: Returns the local host info.
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Property Get LocalHost() As CDnsResult

  Dim strHostName       As String
  Dim lpHost            As Long
  Dim ErrorObject       As CWinsockException

    ' Make room in the buffer
    strHostName = Space(256)

    ' Attempt to get the local host name
    If api_GetHostName(strHostName, 256) = SOCKET_ERROR Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CDns.Get HostName"
        RaiseEvent OnError(ErrorObject)
        Exit Property
    End If
    
    ' Get the hostent info by the name
    ' This will block, but it should block for only a few milliseconds since this is the
    ' local host. So we won't worry about it
    lpHost = api_GetHostByName(strHostName)
    
    ' Return the results
    Set LocalHost = GetResultByHostEntryPointer(lpHost)

End Property


' *************************************************************************************************
' Description: Resolves strHostname either synchronously or asynchronously. strHostname can be an
'              IP address, or a host name such as "www.microsoft.com". If "Async" is false, the
'              function will block and return the resolved information. If "Async" is true, the
'              function will return immediately. The return info will be Nothing and the results
'              returned by the HostResolved event (or if an error occured then the OnError event.
' Author:      Chris Waddell
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function Resolve(strHostName As String, Optional Async As Boolean = False) As CDnsResult

  Dim lpHost As Long
  Dim ErrorObject As CWinsockException
    
    If LenB(strHostName) = 0 Then
        Err.Raise -1, "CDns.Resolve", "The host name cannot be empty"
        Exit Function
    End If
    
    ' If we're already in the middle of resolving a host, don't try to do it again
    If TaskHandle <> 0 Then
        Err.Raise -1, "CDns.Resolve", "You cannot asynchronously resolve more than one host at a time with one instance of CDns. Create a new instance"
        Exit Function
    End If
    
    ' Resolve host synchronously
    If Not Async Then
    
        ' If it's not an IP address, use GetHostByName
        If api_Inet_Addr(strHostName) = INADDR_NONE Then
            lpHost = api_GetHostByName(strHostName)
        
        ' Otherwise use GetHostByAddress
        Else
            lpHost = api_GetHostByAddr(api_Inet_Addr(strHostName), 4, AF_INET)
        End If
        
        Set Resolve = GetResultByHostEntryPointer(lpHost)
    
    ' Resolve host asynchronously
    Else
        
        ' If it's not an IP address, use GetHostByName
        If api_Inet_Addr(strHostName) = INADDR_NONE Then
            TaskHandle = api_WSAAsyncGetHostByName(EventWindow.WindowHandle, ResolveHostMessage, strHostName, ResolveBuffer(0), MAXGETHOSTSTRUCT)
        
        ' Otherwise use GetHostByAddress
        Else
            TaskHandle = api_WSAAsyncGetHostByAddr(EventWindow.WindowHandle, ResolveHostMessage, ByVal VarPtr(api_Inet_Addr(strHostName)), 4, AF_INET, ResolveBuffer(0), MAXGETHOSTSTRUCT)
        End If
        
        ' If the task handle is 0 then the calls failed so raise an error
        If TaskHandle = 0 Then
            Set ErrorObject = New CWinsockException
            ErrorObject.Source = "CDns.Resolve"
            RaiseEvent OnError(ErrorObject)
        End If
    End If

End Function


' *************************************************************************************************
' Description: If there is a host name currently being resolved asynchronously and the client
'              wishes to cancel it then calling this method will do so.
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    08/04/2004 by Chris Waddell
' *************************************************************************************************
Public Sub CancelAsyncResolve()
    If TaskHandle <> 0 Then
        api_WSACancelAsyncRequest (TaskHandle)
        TaskHandle = 0
    End If
End Sub


' *************************************************************************************************
' Description: The GetResultByHostEntryPointer method will take a pointer to an API_HOSTENT
'              structure (passing a pointer to anything else can cause a general protection fault
'              or produce unexpected results. Passing a pointer of 0 will do nothing. The function
'              will then extract the information from the pointer and return a new CDnsResult class
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    08/04/2004 by Chris Waddell
' *************************************************************************************************
Private Function GetResultByHostEntryPointer(lpHost As Long) As CDnsResult

  Dim HostEntry         As API_HOSTENT      ' The hostent structure
  
  Dim lpAddresses       As Long             ' A pointer to the address list
  Dim lngIPAddresses()  As Long             ' The address list
  
  Dim lpAliases         As Long             ' A pointer to the alias list
  Dim strAliasList()    As String           ' The alias list
  
  Dim ListIndex         As Long             ' The current list count (used for IP's and aliases)
  
  Dim NewIP             As CIP4Address
  Dim ErrorObject       As CWinsockException
  Dim result            As CDnsResult

  Dim i                 As Integer

    Set result = New CDnsResult

    ' Just so we don't accidentally cause a protection fault
    If lpHost = 0 Then
        Exit Function
    End If

    ' Copy the the host entry pointed at by lpHost into our local
    ' host entry variable
    RtlMoveMemory ByVal VarPtr(HostEntry), ByVal lpHost, LenB(HostEntry)
    
    ' Get the host name
    result.FullyQualifiedDomainName = StringFromPointer(HostEntry.hName)
    
    ' We only support IP v4
    If HostEntry.hLength = 4 Then
    
        ' The hAddrList field is a pointer to an array of pointers that point to IP addresses.
        ' First copy out the pointer to the pointer list
        RtlMoveMemory ByVal VarPtr(lpAddresses), ByVal HostEntry.hAddrList, 4
        
        ' The list is null terminated so keep going until
        Do Until lpAddresses = 0
            
            ' Resize array
            ReDim Preserve lngIPAddresses(0 To ListIndex)
            
            ' Move the IP address into the array
            RtlMoveMemory ByVal VarPtr(lngIPAddresses(ListIndex)), ByVal lpAddresses, LenB(lngIPAddresses(ListIndex))

            ' Increment the number of IPs
            ListIndex = ListIndex + 1

            ' Get a pointer to the next address
            HostEntry.hAddrList = HostEntry.hAddrList + LenB(HostEntry.hAddrList)
            RtlMoveMemory ByVal VarPtr(lpAddresses), HostEntry.hAddrList, 4
        Loop
        
        ' Create the new address list collection
        Set result.AddressList = New CIP4Collection
        
        ' Loop through each IP address and add it to the address list of the result
        For i = 0 To UBound(lngIPAddresses)
            Set NewIP = New CIP4Address
            NewIP.Address = lngIPAddresses(i)
            result.AddressList.Add NewIP
            Set NewIP = Nothing
        Next i
    Else
        Set NewIP = New CIP4Address
        NewIP.StringAddress = "0.0.0.0"
        result.AddressList.Add NewIP
    End If
    
    ListIndex = -1

    ' The hAliases field is a pointer to an array of pointers that point to strings.
    ' First copy out the pointer to the pointer list
    Set result.Aliases = PtrArrayToStrCollection(HostEntry.hAliases)

    Set GetResultByHostEntryPointer = result

End Function


' *************************************************************************************************
' Description: Occurs when a new message arrives at the wndproc.
' Copyright:   Copyright (c) 2004 Chris Waddell
' Contact:     IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:    08/04/2004 by Chris Waddell
' *************************************************************************************************
Private Sub EventWindow_MessageArrived(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, lngRetVal As Long, bHandled As Boolean)

  Dim lpHost As Long
  Dim lngErrorCode As Long
  Dim ErrorObject As CWinsockException

    ' If this isn't anything to do with our host resolving message then just ignore it
    ' and let it be handled in the default way
    If uMsg <> ResolveHostMessage Then
        bHandled = False
        Exit Sub
    End If
    
    ' Set the task handle back to 0 so we know we can use this to resolve again
    TaskHandle = 0
    
    ' Extract the error code from the hi word of lParam
    lngErrorCode = HiWord(lParam)
    
    ' If there is an error then raise the OnError event, otherwise extract the results and return
    ' then via the HostResolved event
    If lngErrorCode <> 0 Then
        Set ErrorObject = New CWinsockException
        ErrorObject.Source = "CDns.HostResolved"
        ErrorObject.ErrorCode = lngErrorCode
        RaiseEvent OnError(ErrorObject)
    Else
        RaiseEvent HostResolved(GetResultByHostEntryPointer(VarPtr(ResolveBuffer(0))))
    End If

    ' We've handled this
    lngRetVal = 0
    bHandled = True

End Sub


Private Sub Class_Initialize()
    ResolveHostMessage = RegisterWindowMessage(App.EXEName & ":" & App.hInstance & ":Resolver")
    Set EventWindow = New CWindow
End Sub


Private Sub Class_Terminate()
    CancelAsyncResolve
    Set EventWindow = Nothing
End Sub
