Attribute VB_Name = "modWinsockUtils"
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


'String functions
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

' Memory functions
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDst As Long, ByVal pSrc As Long, ByVal ByteLen As Long)

' GUID functions
Public Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpsz As String, ByVal cchMax As Long) As Long
                 

' *************************************************************************************************
' Description:  The format of the description of the error code
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Enum ErrorCodeDescriptionType

    ' A short description of the error
    DescriptionShort
    
    ' A long detailed description of the error
    DescriptionLong
    
    ' The win32 API constant name of the error
    DescriptionWin32API
    
End Enum


' The number of times the windows sockets service has been started by this project
Public WSAStartupCount As Long


' *************************************************************************************************
' Description:  Use this macro to make the winsock version.
'               eg m_intVersion = MakeVersion(&h2, &h2)
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function MakeVersion(Major As Byte, Minor As Byte) As Integer
    MakeVersion = MakeWord(Major, Minor)
End Function


' *************************************************************************************************
' Description:  Use this macro to get the major part from a winsock version.
'               eg m_bytMajor = GetMajor(&h202)
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetMajor(Version As Integer) As Byte
    GetMajor = HiByte(Version)
End Function


' *************************************************************************************************
' Description:  Use this macro to get the minor part from a winsock version.
'               eg m_bytMinor = GetMinor(&h202)
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetMinor(Version As Integer) As Byte
    GetMinor = LoByte(Version)
End Function


' *************************************************************************************************
' Description:  Use this macro to pack 2 bytes into a word
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function MakeWord(Hi As Byte, Lo As Byte) As Integer
  Dim Word(0 To 1) As Byte
    
    Let Word(0) = Lo
    Let Word(1) = Hi
    
    Call RtlMoveMemory(ByVal VarPtr(MakeWord), ByVal VarPtr(Word(0)), 2)
End Function


' *************************************************************************************************
' Description:  Use this macro to return the low byte from a 16 bit word
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function LoByte(Value As Integer) As Byte
    Call RtlMoveMemory(ByVal VarPtr(LoByte), ByVal VarPtr(Value), 1)
End Function


' *************************************************************************************************
' Description:  Use this macro to return the high byte from a 16 bit word
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function HiByte(Value As Integer) As Byte
    Call RtlMoveMemory(ByVal VarPtr(HiByte), ByVal VarPtr(Value) + 1, 1)
End Function


' *************************************************************************************************
' Description:  Use this macro to return the low 16 bit word from a 32 bit word
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function LoWord(Value As Long) As Integer
    Call RtlMoveMemory(ByVal VarPtr(LoWord), ByVal VarPtr(Value), 2)
End Function


' *************************************************************************************************
' Description:  Use this macro to return the high 16 bit word from a 32 bit word
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function HiWord(Value As Long) As Integer
    Call RtlMoveMemory(ByVal VarPtr(HiWord), ByVal VarPtr(Value) + 2, 2)
End Function


' *************************************************************************************************
' Description:  A function to convert a GUID structure to it's string equivelant
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function GuidToString(Guid As API_GUID) As String

  Dim strGuid As String
  Dim lngRetVal As Long

    strGuid = String(80, Chr$(0))
    lngRetVal = StringFromGUID2(Guid, strGuid, 80&)
    
    If lngRetVal > 0 Then GuidToString = StrConv(strGuid, vbFromUnicode)

End Function


' *************************************************************************************************
' Description:  Returns a description of the error code
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetErrorDescription(ByVal ErrorCode As ErrorCodeType, Optional ByVal DescriptionType As ErrorCodeDescriptionType = ErrorCodeDescriptionType.DescriptionShort) As String


  Dim ShortDescription  As String
  Dim LongDescription   As String
  Dim Win32APIName      As String


    Select Case ErrorCode

        Case ErrorCodeType.NoError
            
            ShortDescription = "No Error"
            LongDescription = "The operation completed successfuly"
            Win32APIName = "NO_ERROR"
            
        Case ErrorCodeType.InterruptedFunctionCall
    
            ShortDescription = "Interrupted function call"
            LongDescription = "A blocking operation was interrupted by a call to WSACancelBlockingCall."
            Win32APIName = "WSAEINTR"
            
        Case ErrorCodeType.PermissionDenied
            
            ShortDescription = "Permission denied"
            LongDescription = "An attempt was made to access a socket in a way forbidden by its access permissions. " & _
                              "An example is using a broadcast address for sendto without broadcast permission being " & _
                              "set using setsockopt(SO_BROADCAST)." & vbCrLf & vbCrLf & _
                              "Another possible reason for the WSAEACCES error is that when the bind function is called " & _
                              "(on Windows NT 4 SP4 or later), another application, service, or kernel mode driver is " & _
                              "bound to the same address with exclusive access. Such exclusive access is a new feature " & _
                              "of Windows NT 4 SP4 and later, and is implemented by using the SO_EXCLUSIVEADDRUSE option."
            Win32APIName = "WSAEACCES"
            
        Case ErrorCodeType.BadAddress
        
            ShortDescription = "Bad address"
            LongDescription = "The system detected an invalid pointer address in attempting to use a pointer argument of " & _
                              "a call. This error occurs if an application passes an invalid pointer value, or if the " & _
                              "length of the buffer is too small. For instance, if the length of an argument, which is a " & _
                              "sockaddr structure, is smaller than the sizeof(sockaddr)."
            Win32APIName = "WSAEFAULT"
        
        Case ErrorCodeType.InvalidArgument
            
            ShortDescription = "Invalid argument"
            LongDescription = "Some invalid argument was supplied (for example, specifying an invalid level to the " & _
                              "setsockopt function). In some instances, it also refers to the current state of the " & _
                              "socket—for instance, calling accept on a socket that is not listening."
            Win32APIName = "WSAEINVAL"
        
        Case ErrorCodeType.TooManyOpenFiles
    
            ShortDescription = "Too many open files"
            LongDescription = "Each implementation may have a maximum number of socket handles available, either " & _
            "globally, per process, or per thread"
            Win32APIName = "WSAEMFILE"
    
        Case ErrorCodeType.ResourceTemporarilyUnavailable
            
            ShortDescription = "Resource temporarily unavailable"
            LongDescription = "This error is returned from operations on nonblocking sockets that cannot be completed " & _
                              "immediately, for example recv when no data is queued to be read from the socket. It is a " & _
                              "nonfatal error, and the operation should be retried later. It is normal for WSAEWOULDBLOCK " & _
                              "to be reported as the result from calling connect on a nonblocking SOCK_STREAM socket, " & _
                              "since some time must elapse for the connection to be established."
            Win32APIName = "WSAEWOULDBLOCK"
        
        Case ErrorCodeType.OperationNowInProgress
        
            ShortDescription = "Operation now in progress"
            LongDescription = "A blocking operation is currently executing. Windows Sockets only allows a single blocking " & _
                              "operation—per- task or thread—to be outstanding, and if any other function call is made " & _
                              "(whether or not it references that or any other socket) the function fails with the " & _
                              "WSAEINPROGRESS error."
            Win32APIName = "WSAEINPROGRESS"
        
        Case ErrorCodeType.OperationAlreadyInProgress
        
            ShortDescription = "Operation already in progress"
            LongDescription = "An operation was attempted on a nonblocking socket with an operation already in " & _
                              "progress—that is, calling connect a second time on a nonblocking socket that is already " & _
                              "connecting, or canceling an asynchronous request (WSAAsyncGetXbyY) that has already been " & _
                              "canceled or completed."
            Win32APIName = "WSAEALREADY"
        
        Case ErrorCodeType.SocketOperationOnNonSocket
        
            ShortDescription = "Socket operation on non socket"
            LongDescription = "An operation was attempted on something that is not a socket. Either the socket handle " & _
                              "parameter did not reference a valid socket, or for select, a member of an fd_set was not " & _
                              "valid."
            Win32APIName = "WSAENOTSOCK"
            
        Case ErrorCodeType.DestinationAddressRequired
    
            ShortDescription = "Destination address required"
            LongDescription = "A required address was omitted from an operation on a socket. For example, this error is " & _
                              "returned if sendto is called with the remote address of ADDR_ANY."
            Win32APIName = "WSAEDESTADDRREQ"
            
        Case ErrorCodeType.MessageTooLong
        
            ShortDescription = "Message too long"
            LongDescription = "A message sent on a datagram socket was larger than the internal message buffer or some " & _
                              "other network limit, or the buffer used to receive a datagram was smaller than the " & _
                              "datagram itself."
            Win32APIName = "WSAEMSGSIZE"
    
        Case ErrorCodeType.ProtocolWrongTypeForSocket
        
            ShortDescription = "Protocol wrong type for socket"
            LongDescription = "A protocol was specified in the socket function call that does not support the semantics " & _
                              "of the socket type requested. For example, the ARPA Internet UDP protocol cannot be " & _
                              "specified with a socket type of SOCK_STREAM."
            Win32APIName = "WSAEPROTOTYPE"
        
        Case ErrorCodeType.BadProtocolOption
        
            ShortDescription = "Bad protocol option"
            LongDescription = "An unknown, invalid or unsupported option or level was specified in a getsockopt or " & _
                              "setsockopt call."
            Win32APIName = "WSAENOPROTOOPT"
        
        Case ErrorCodeType.ProtocolNotSupported
    
            ShortDescription = "Protocol not supported"
            LongDescription = "The requested protocol has not been configured into the system, or no implementation for " & _
                              "it exists. For example, a socket call requests a SOCK_DGRAM socket, but specifies a stream " & _
                              "protocol."
            Win32APIName = "WSAEPROTONOSUPPORT"
    
        Case ErrorCodeType.SocketTypeNotSupported
    
            ShortDescription = "Socket type not supported"
            LongDescription = "The support for the specified socket type does not exist in this address family. " & _
                              "For example, the optional type SOCK_RAW might be selected in a socket call, and the " & _
                              "implementation does not support SOCK_RAW sockets at all."
            Win32APIName = "WSAESOCKTNOSUPPORT"
    
        Case ErrorCodeType.OperationNotSupported
    
            ShortDescription = "Operation not supported"
            LongDescription = "The attempted operation is not supported for the type of object referenced. Usually this " & _
                              "occurs when a socket descriptor to a socket that cannot support this operation is trying " & _
                              "to accept a connection on a datagram socket."
            Win32APIName = "WSAEOPNOTSUPP"
    
        Case ErrorCodeType.ProtocolFamilyNotSupported
    
            ShortDescription = "Protocol family not supported"
            LongDescription = "The protocol family has not been configured into the system or no implementation for it " & _
                              "exists. This message has a slightly different meaning from WSAEAFNOSUPPORT. However, it " & _
                              "is interchangeable in most cases, and all Windows Sockets functions that return one of " & _
                              "these messages also specify WSAEAFNOSUPPORT."
            Win32APIName = "WSAEPFNOSUPPORT"
    
        Case ErrorCodeType.AddressFamilyNotSupportedByProtocolFamily
    
            ShortDescription = "Address family not supported by protocol family"
            LongDescription = "An address incompatible with the requested protocol was used. All sockets are created with " & _
                              "an associated address family (that is, AF_INET for Internet Protocols) and a generic " & _
                              "protocol type (that is, SOCK_STREAM). This error is returned if an incorrect protocol is " & _
                              "explicitly requested in the socket call, or if an address of the wrong family is used for " & _
                              "a socket, for example, in sendto."
            Win32APIName = "WSAEAFNOSUPPORT"
    
        Case ErrorCodeType.AddressAlreadyInUse
    
            ShortDescription = "Address already in use"
            LongDescription = "Typically, only one usage of each socket address (protocol/IP address/port) is permitted. " & _
                              "This error occurs if an application attempts to bind a socket to an IP address/port that " & _
                              "has already been used for an existing socket, or a socket that was not closed properly, or " & _
                              "one that is still in the process of closing. For server applications that need to bind " & _
                              "multiple sockets to the same port number, consider using setsockopt (SO_REUSEADDR). " & _
                              "Client applications usually need not call bind at all—connect chooses an unused port " & _
                              "automatically. When bind is called with a wildcard address (involving ADDR_ANY), a " & _
                              "WSAEADDRINUSE error could be delayed until the specific address is committed. This could " & _
                              "happen with a call to another function later, including connect, listen, WSAConnect, or " & _
                              "WSAJoinLeaf."
            Win32APIName = "WSAEADDRINUSE"
    
        Case ErrorCodeType.CannotAssignRequestedAddress
    
            ShortDescription = "Cannot assign requested address"
            LongDescription = "The requested address is not valid in its context. This normally results from an attempt to " & _
                              "bind to an address that is not valid for the local computer. This can also result from " & _
                              "connect, sendto, WSAConnect, WSAJoinLeaf, or WSASendTo when the remote address or port is " & _
                              "not valid for a remote computer (for example, address or port 0)."
            Win32APIName = "WSAEADDRNOTAVAIL"
    
        Case ErrorCodeType.NetworkIsDown
    
            ShortDescription = "Network os down"
            LongDescription = "A socket operation encountered a dead network. This could indicate a serious failure of the " & _
                              "network system (that is, the protocol stack that the Windows Sockets DLL runs over), the " & _
                              "network interface, or the local network itself."
            Win32APIName = "WSAENETDOWN"
    
        Case ErrorCodeType.NetworkIsUnreachable
    
            ShortDescription = "Network is unreachable"
            LongDescription = "A socket operation was attempted to an unreachable network. This usually means the local " & _
                              "software knows no route to reach the remote host."
            Win32APIName = "WSAENETUNREACH"
    
        Case ErrorCodeType.NetworkDroppedConnectionOnReset
    
            ShortDescription = "Network dropped connection on reset"
            LongDescription = "The connection has been broken due to keep-alive activity detecting a failure while the " & _
                              "operation was in progress. It can also be returned by setsockopt if an attempt is made to " & _
                              "set SO_KEEPALIVE on a connection that has already failed."
            Win32APIName = "WSAENETRESET"

        Case ErrorCodeType.SoftwareCausedConnectionAbort
    
            ShortDescription = "Software caused connection abort"
            LongDescription = "An established connection was aborted by the software in your host computer, possibly due " & _
                              "to a data transmission time-out or protocol error."
            Win32APIName = "WSAECONNABORTED"
    
        Case ErrorCodeType.ConnectionResetByPeer
    
            ShortDescription = "Connection reset by peer"
            LongDescription = "An existing connection was forcibly closed by the remote host. This normally results if the " & _
                              "peer application on the remote host is suddenly stopped, the host is rebooted, the host or " & _
                              "remote network interface is disabled, or the remote host uses a hard close (see setsockopt " & _
                              "for more information on the SO_LINGER option on the remote socket). This error may also " & _
                              "result if a connection was broken due to keep-alive activity detecting a failure while one " & _
                              "or more operations are in progress. Operations that were in progress fail with WSAENETRESET. " & _
                              "Subsequent operations fail with WSAECONNRESET."
            Win32APIName = "WSAECONNRESET"
    
        Case NoBufferSpaceAvailable
    
            ShortDescription = "No buffer space available"
            LongDescription = "An operation on a socket could not be performed because the system lacked sufficient buffer " & _
                              "space or because a queue was full."
            Win32APIName = "WSAENOBUFS"

        Case ErrorCodeType.SocketIsAlreadyConnected
    
            ShortDescription = "Socket is already connected"
            LongDescription = "A connect request was made on an already-connected socket. Some implementations also " & _
                              "return this error if sendto is called on a connected SOCK_DGRAM socket (for SOCK_STREAM " & _
                              "sockets, the to parameter in sendto is ignored) although other implementations treat this " & _
                              "as a legal occurrence."
            Win32APIName = "WSAEISCONN"
    
        Case ErrorCodeType.SocketIsNotConnected
    
            ShortDescription = "Socket is not connected"
            LongDescription = "A request to send or receive data was disallowed because the socket is not connected and " & _
                              "(when sending on a datagram socket using sendto) no address was supplied. Any other type " & _
                              "of operation might also return this error—for example, setsockopt setting SO_KEEPALIVE if " & _
                              "the connection has been reset."
            Win32APIName = "WSAENOTCONN"

        Case ErrorCodeType.CannotSendAfterSocketShutdown
    
            ShortDescription = "Cannot send after socket shutdown"
            LongDescription = "A request to send or receive data was disallowed because the socket had already been shut " & _
                              "down in that direction with a previous shutdown call. By calling shutdown a partial close " & _
                              "of a socket is requested, which is a signal that sending or receiving, or both have been " & _
                              "discontinued."
            Win32APIName = "WSAESHUTDOWN"

        Case ErrorCodeType.ConnectionTimedOut
    
            ShortDescription = "Connection timed out"
            LongDescription = "A connection attempt failed because the connected party did not properly respond after a " & _
                              "period of time, or the established connection failed because the connected host has failed " & _
                              "to respond."
            Win32APIName = "WSAETIMEDOUT"
    
        Case ErrorCodeType.ConnectionRefused
    
            ShortDescription = "Connection refused"
            LongDescription = "No connection could be made because the target computer actively refused it. This usually " & _
                              "results from trying to connect to a service that is inactive on the foreign host—that is, " & _
                              "one with no server application running."
            Win32APIName = "WSAECONNREFUSED"
    
        Case ErrorCodeType.HostIsDown
    
            ShortDescription = "Host is down"
            LongDescription = "A socket operation failed because the destination host is down. A socket operation " & _
                              "encountered a dead host. Networking activity on the local host has not been initiated. " & _
                              "These conditions are more likely to be indicated by"
            Win32APIName = "WSAEHOSTDOWN"

        Case ErrorCodeType.NoRouteToHost
    
            ShortDescription = "No route to host"
            LongDescription = "A socket operation was attempted to an unreachable host. See WSAENETUNREACH. "
            Win32APIName = "WSAEHOSTUNREACH"

        Case ErrorCodeType.TooManyProcesses
    
            ShortDescription = "Too many processes"
            LongDescription = "A Windows Sockets implementation may have a limit on the number of applications that can " & _
                              "use it simultaneously. WSAStartup may fail with this error if the limit has been reached."
            Win32APIName = "WSAEPROCLIM"
    
        Case ErrorCodeType.NetworkSubsystemIsUnavailable
    
            ShortDescription = "Network subsystem is unavailable"
            LongDescription = "This error is returned by WSAStartup if the Windows Sockets implementation cannot function " & _
                              "at this time because the underlying system it uses to provide network services is currently " & _
                              "unavailable. Users should check: " & vbCrLf & vbCrLf & _
                              "- That the appropriate Windows Sockets DLL file is in the current path. " & vbCrLf & vbCrLf & _
                              "- That they are not trying to use more than one Windows Sockets implementation " & _
                              "  simultaneously. If there is more than one Winsock DLL on your system, be sure the first " & _
                              "  one in the path is appropriate for the network subsystem currently loaded. " & vbCrLf & vbCrLf & _
                              "- The Windows Sockets implementation documentation to be sure all necessary components are " & _
                              "  currently installed and configured correctly. "
            Win32APIName = "WSASYSNOTREADY"
    
        Case ErrorCodeType.WinsockDllVersionOutOfRange
    
            ShortDescription = "Winsock.dll version out of range"
            LongDescription = "The current Windows Sockets implementation does not support the Windows Sockets " & _
                              "specification version requested by the application. Check that no old Windows Sockets DLL " & _
                              "files are being accessed."
            Win32APIName = "WSAVERNOTSUPPORTED"
    
        Case ErrorCodeType.SuccessfulWSAStartupNotYetPerformed
    
            ShortDescription = "Successful WSAStartup not yet performed"
            LongDescription = "Either the application has not called WSAStartup or WSAStartup failed. The application may " & _
                              "be accessing a socket that the current active task does not own (that is, trying to share a " & _
                              "socket between tasks), or WSACleanup has been called too many times."
            Win32APIName = "WSANOTINITIALISED"
    
        Case ErrorCodeType.GracefulShutdownInProgress
    
            ShortDescription = "Graceful shutdown in progress"
            LongDescription = "Returned by WSARecv and WSARecvFrom to indicate that the remote party has initiated a " & _
                              "graceful shutdown sequence."
            Win32APIName = "WSAEDISCON"
    
        Case ErrorCodeType.ClassTypeNotFound
    
            ShortDescription = "Class type not found"
            LongDescription = "The specified class was not found."
            Win32APIName = "WSATYPE_NOT_FOUND"

        Case ErrorCodeType.HostNotFound
        
            ShortDescription = "Host not found"
            LongDescription = "No such host is known. The name is not an official host name or alias, or it cannot be " & _
                              "found in the database(s) being queried. This error may also be returned for protocol and " & _
                              "service queries, and means that the specified name could not be found in the relevant " & _
                              "database."
            Win32APIName = "WSAHOST_NOT_FOUND"
        
        Case ErrorCodeType.NonAuthoritativeHostNotFound
    
            ShortDescription = "Non authoritative host not found"
            LongDescription = "This is usually a temporary error during host name resolution and means that the local " & _
                              "server did not receive a response from an authoritative server. A retry at some time later " & _
                              "may be successful."
            Win32APIName = "WSATRY_AGAIN"
    
        Case ErrorCodeType.NonRecoverableError
    
            ShortDescription = "Non recoverable error"
            LongDescription = "This indicates some sort of nonrecoverable error occurred during a database lookup. This " & _
                              "may be because the database files (for example, BSD-compatible HOSTS, SERVICES, or " & _
                              "PROTOCOLS files) could not be found, or a DNS request was returned by the server with a " & _
                              "severe error."
            Win32APIName = "WSANO_RECOVERY"
    
        
        Case ErrorCodeType.ValidNameButNoDataRecordOfRequestedType
    
            ShortDescription = "Valid name but no data record of requested type"
            LongDescription = "The requested name is valid and was found in the database, but it does not have the " & _
                              "correct associated data being resolved for. The usual example for this is a host " & _
                              "name-to-address translation attempt (using gethostbyname or WSAAsyncGetHostByName) " & _
                              "which uses the DNS (Domain Name Server). An MX record is returned but no A " & _
                              "record—indicating the host itself exists, but is not directly reachable."
            Win32APIName = "WSANO_DATA"
    
        
        Case ErrorCodeType.InvalidHandle
    
            ShortDescription = "Invalid handle"
            LongDescription = "Specified event object handle is invalid. " & vbCrLf & _
                              "An application attempts to use an event object, but the specified handle is not valid. "
            Win32APIName = "WSA_INVALID_HANDLE"
    
        
        Case ErrorCodeType.InvalidParameter
    
            ShortDescription = "Invalid parameter"
            LongDescription = "An application used a Windows Sockets function which directly maps to a Windows function. " & _
                              "The Windows function is indicating a problem with one or more parameters."
            Win32APIName = "WSA_INVALID_PARAMETER"
    
        
        Case ErrorCodeType.IOEventIncomplete
    
            ShortDescription = "IO event incomplete"
            LongDescription = "Overlapped I/O event object not in signaled state. " & vbCrLf & _
                              "The application has tried to determine the status of an overlapped operation which is not " & _
                              "yet completed. Applications that use WSAGetOverlappedResult (with the fWait flag set to " & _
                              "FALSE) in a polling mode to determine when an overlapped operation has completed."
            Win32APIName = "WSA_IO_INCOMPLETE"
    
        
        Case ErrorCodeType.IOOperationPending
        
            ShortDescription = "IO operation pending"
            LongDescription = "Overlapped operations will complete later. " & vbCrLf & _
                              "The application has initiated an overlapped operation that cannot be completed immediately. " & _
                              "A completion indication will be given later when the operation has been completed."
            Win32APIName = "WSA_IO_PENDING"
        
        
        Case ErrorCodeType.InsufficientMemory
        
            ShortDescription = "Insufficient memory"
            LongDescription = "Insufficient memory available." & vbCrLf & _
                              "An application used a Windows Sockets function that directly maps to a Windows function. " & _
                              "The Windows function is indicating a lack of required memory resources."
            Win32APIName = "WSA_NOT_ENOUGH_MEMORY"
        
        
        Case ErrorCodeType.OverlappedOperationAborted
        
            ShortDescription = "Overlapped operation aborted"
            LongDescription = "Overlapped operation aborted." & vbCrLf & _
                              "An overlapped operation was canceled due to the closure of the socket, or the execution of " & _
                              "the SIO_FLUSH command in WSAIoctl."
            Win32APIName = "WSA_OPERATION_ABORTED"
        
        
        Case ErrorCodeType.InvalidProcedureTable
        
            ShortDescription = "Invalid procedure table"
            LongDescription = "Invalid procedure table from service provider. " & vbCrLf & _
                              "A service provider returned a bogus procedure table to Ws2_32.dll. (Usually caused by one  " & _
                              "or more of the function pointers being null.)"
            Win32APIName = "WSAINVALIDPROCTABLE"
        
        
        Case ErrorCodeType.InvalidServiceProviderVersion
        
            ShortDescription = "Invalid service provider version"
            LongDescription = "Invalid service provider version number. " & vbCrLf & _
                              "A service provider returned a version number other than 2.0."
            Win32APIName = "WSAINVALIDPROVIDER"
        
        
        Case ErrorCodeType.UnableToInitializeServiceProvider
        
            ShortDescription = "Unable to initialize service provider"
            LongDescription = "Unable to initialize a service provider. " & vbCrLf & _
                              "Either a service provider's DLL could not be loaded (LoadLibrary failed) or the provider's  " & _
                              "WSPStartup/NSPStartup function failed."
            Win32APIName = "WSAPROVIDERFAILEDINIT"
        
       
        Case ErrorCodeType.SystemCallFailure

            ShortDescription = "System call failure"
            LongDescription = "System call failure. " & vbCrLf & _
                              "Generic error code, returned under various conditions.  " & vbCrLf & vbCrLf & _
                              "- Returned when a system call that should never fail does fail. For example, if a call to  " & _
                              "  WaitForMultipleEvents fails or one of the registry functions fails trying to manipulate  " & _
                              "  the protocol/namespace catalogs." & vbCrLf & vbCrLf & _
                              "- Returned when a provider does not return SUCCESS and does not provide an extended error  " & _
                              "  code. Can indicate a service provider implementation error."
            Win32APIName = "WSASYSCALLFAILURE"
        
        Case Else
        
            ShortDescription = "Unknown error code"
            LongDescription = "Unknown error code"
            Win32APIName = "UNKNOWN"
        
    End Select
    
    
    Select Case DescriptionType
        Case ErrorCodeDescriptionType.DescriptionWin32API
            GetErrorDescription = Win32APIName
        Case ErrorCodeDescriptionType.DescriptionShort
            GetErrorDescription = ShortDescription
        Case ErrorCodeDescriptionType.DescriptionLong
            GetErrorDescription = LongDescription
        Case Else
            GetErrorDescription = ShortDescription
    End Select
    
End Function


' *************************************************************************************************
' Description:  Convert a long representing an unsigned 16 bit number into a signed integer
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function UnsignedToInteger(Value As Long) As Integer

    If Value < 0 Or Value >= 2 ^ 16 Then Error 6 'Overflow
    
    If Value <= (2 ^ 16) / 2 - 1 Then
        Let UnsignedToInteger = Value
    Else
        Let UnsignedToInteger = Value - 2 ^ 16
    End If

End Function


' *************************************************************************************************
' Description:  Convert an integer containing a signed 16 bit number into a long, representing the
'               unsigned version
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function IntegerToUnsigned(Value As Integer) As Long

    If Value < 0 Then
        Let IntegerToUnsigned = Value + 2 ^ 16
    Else
        Let IntegerToUnsigned = Value
    End If
    
End Function


' *************************************************************************************************
' Description:  Convert a currency representing an unsigned 32 bit number into a signed long
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function UnsignedToLong(Value As Currency) As Long

    If Value < 0 Or Value >= 2 ^ 32 Then Error 6 'Overflow
    
    If Value <= (2 ^ 32) / 2 - 1 Then
        Let UnsignedToLong = Value
    Else
        Let UnsignedToLong = Value - 2 ^ 32
    End If

End Function


' *************************************************************************************************
' Description:  Convert a long containing a signed 32 bit number into a currency, representing the
'               unsigned version
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function LongToUnsigned(Value As Long) As Currency

    If Value < 0 Then
        Let LongToUnsigned = Value + 2 ^ 32
    Else
        Let LongToUnsigned = Value
    End If
    
End Function



' *************************************************************************************************
' Description:  Get a string by it's pointer
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function StringFromPointer(ByVal lngPointer As Long) As String

  Dim strTemp As String
  Dim lRetVal As Long
    
    strTemp = String$(lstrlen(ByVal lngPointer), 0)    'prepare the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lngPointer) 'copy the string into the strTemp buffer
    If lRetVal Then StringFromPointer = strTemp        'return the string

End Function


' *************************************************************************************************
' Description:  Convert a long IP to it's string representation
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetStrIPFromLong(lngIP As Long) As String

  Dim lpstrAddress As Long
   
    lpstrAddress = modWinsock.api_Inet_ntoa(lngIP)
   
    If lpstrAddress = 0 Then
        GetStrIPFromLong = "0.0.0.0"
    Else
        GetStrIPFromLong = StringFromPointer(lpstrAddress)
    End If

End Function


' *************************************************************************************************
' Description:  Convert a string IP to it's long representation
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetLngIPFromStr(strIP As String) As Long
    GetLngIPFromStr = modWinsock.api_Inet_Addr(strIP)
End Function


' *************************************************************************************************
' Description:  Get a string collection from a null terminated array of string pointers
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function PtrArrayToStrCollection(ByVal lngPointer As Long) As Collection

    Dim lpString As Long
    Dim colList As Collection

    If lngPointer <> 0 Then
        Set colList = New Collection

        'Get pointer to the first item in the list
        RtlMoveMemory ByVal VarPtr(lpString), ByVal lngPointer, 4&

        ' The list is null terminated so keep going until
        Do Until lpString = 0
        
            ' Copy the string into the collection
            colList.Add StringFromPointer(lpString)

            ' Get a pointer to the next string
            lngPointer = lngPointer + 4
            RtlMoveMemory ByVal VarPtr(lpString), ByVal lngPointer, 4&
        Loop

        Set PtrArrayToStrCollection = colList

    End If

End Function


' *************************************************************************************************
' Description:  The GetAddressFamilyName function returns a human readable string representation of
'               an address family.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     06/04/2004 by Chris Waddellrg
' Modified:     28/03/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetAddressFamilyName(AddressFamily As AddressFamilyType) As String

    Select Case AddressFamily
        Case AF_UNSPEC:        GetAddressFamilyName = "Unspecified"
        Case AF_UNIX:          GetAddressFamilyName = "Unix"
        Case AF_INET:          GetAddressFamilyName = "Internetwork"
        Case AF_IMPLINK:       GetAddressFamilyName = "ImpLink"
        Case AF_PUP:           GetAddressFamilyName = "PUP"
        Case AF_CHAOS:         GetAddressFamilyName = "Chaos"
        Case AF_NS:            GetAddressFamilyName = "Xerox NS"
        Case AF_IPX:           GetAddressFamilyName = "IPX"
        Case AF_ISO:           GetAddressFamilyName = "ISO"
        Case AF_OSI:           GetAddressFamilyName = "ISO"
        Case AF_ECMA:          GetAddressFamilyName = "ECMA"
        Case AF_DATAKIT:       GetAddressFamilyName = "Datakit"
        Case AF_CCITT:         GetAddressFamilyName = "CCITT"
        Case AF_SNA:           GetAddressFamilyName = "IBM SNA"
        Case AF_DECnet:        GetAddressFamilyName = "DECnet"
        Case AF_DLI:           GetAddressFamilyName = "Direct data link interface"
        Case AF_LAT:           GetAddressFamilyName = "LAT"
        Case AF_HYLINK:        GetAddressFamilyName = "NSC Hyperchannel"
        Case AF_APPLETALK:     GetAddressFamilyName = "AppleTalk"
        Case AF_NETBIOS:       GetAddressFamilyName = "NetBios"
        Case AF_VOICEVIEW:     GetAddressFamilyName = "VoiceView"
        Case AF_FIREFOX:       GetAddressFamilyName = "Firefox"
        Case AF_BAN:           GetAddressFamilyName = "Banyan"
        Case AF_ATM:           GetAddressFamilyName = "Native ATM Services"
        Case AF_INET6:         GetAddressFamilyName = "Internetwork Version 6"
        Case AF_CLUSTER:       GetAddressFamilyName = "Microsoft Wolfpack"
        Case AF_12844:         GetAddressFamilyName = "IEEE 1284.4 WG AF"
        Case AF_IRDA:          GetAddressFamilyName = "IrDA"
        Case AF_NETDES:        GetAddressFamilyName = "Network Designer"
        Case Else:             GetAddressFamilyName = "Unknown Address Family"
    End Select

End Function


' *************************************************************************************************
' Description:  The GetSocketTypeName function returns a human readable string representation of
'               a socket type.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     06/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetSocketTypeName(SockType As SocketType) As String
    
    Select Case SockType
        Case SOCK_STREAM:    GetSocketTypeName = "Stream Socket"
        Case SOCK_DGRAM:     GetSocketTypeName = "Datagram Socket"
        Case SOCK_RAW:       GetSocketTypeName = "Raw Protocol Interface"
        Case SOCK_RDM:       GetSocketTypeName = "Reliably Delivered Message"
        Case SOCK_SEQPACKET: GetSocketTypeName = "Sequenced Packet Stream"
        Case Else:           GetSocketTypeName = "Unknown Socket Type"
    End Select
    
End Function


' *************************************************************************************************
' Description:  The GetProtocolName function returns a human readable string representation of
'               a protocol.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     06/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetProtocolName(Protocol As IPProtocolType) As String

    Select Case Protocol
        Case IPPROTO_IP:    GetProtocolName = "IP (Internet Protocol)"
        Case IPPROTO_ICMP:  GetProtocolName = "ICMP (Internet Control Message Protocol)"
        Case IPPROTO_IGMP:  GetProtocolName = "IGMP (Internet Group Management Protocol)"
        Case IPPROTO_GGP:   GetProtocolName = "GGP (Gateway^2)"
        Case IPPROTO_TCP:   GetProtocolName = "TCP (Transmission Control Protocol)"
        Case IPPROTO_PUP:   GetProtocolName = "PUP"
        Case IPPROTO_UDP:   GetProtocolName = "UDP (User Datagram Protocol)"
        Case IPPROTO_IDP:   GetProtocolName = "IDP"
        Case IPPROTO_ND:    GetProtocolName = "ND (NetDisk)"
        Case IPPROTO_IPX:   GetProtocolName = "IPX"
        Case IPPROTO_SPX:   GetProtocolName = "SPX"
        Case IPPROTO_SPXII: GetProtocolName = "SPX II"
        Case IPPROTO_RAW:   GetProtocolName = "Raw IP"
        Case Else:          GetProtocolName = "Unknown Protocol"
    End Select

End Function


' *************************************************************************************************
' Description:  Gets a byte array containing the data passed in value.
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetBufferByVariant(varData As Variant) As Byte()

  Dim arrData() As Byte

    Select Case VarType(varData)
        Case vbArray Or vbByte
            arrData() = varData
        Case vbBoolean
            Dim blnData As Boolean
            blnData = CBool(varData)
            ReDim arrData(LenB(blnData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(blnData), LenB(blnData)
        Case vbByte
            Dim bytData As Byte
            bytData = CByte(varData)
            ReDim arrData(LenB(bytData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(bytData), LenB(bytData)
        Case vbCurrency
            Dim curData As Currency
            curData = CCur(varData)
            ReDim arrData(LenB(curData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(curData), LenB(curData)
        Case vbDate
            Dim datData As Date
            datData = CDate(varData)
            ReDim arrData(LenB(datData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(datData), LenB(datData)
        Case vbDouble
            Dim dblData As Double
            dblData = CDbl(varData)
            ReDim arrData(LenB(dblData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(dblData), LenB(dblData)
        Case vbInteger
            Dim intData As Integer
            intData = CInt(varData)
            ReDim arrData(LenB(intData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(intData), LenB(intData)
        Case vbLong
            Dim lngData As Long
            lngData = CLng(varData)
            ReDim arrData(LenB(lngData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(lngData), LenB(lngData)
        Case vbSingle
            Dim sngData As Single
            sngData = CSng(varData)
            ReDim arrData(LenB(sngData) - 1)
            RtlMoveMemory ByVal VarPtr(arrData(0)), ByVal VarPtr(sngData), LenB(sngData)
        Case vbString
            Dim strData As String
            strData = CStr(varData)
            ReDim arrData(Len(strData) - 1)
            arrData() = StrConv(strData, vbFromUnicode)
    End Select

    GetBufferByVariant = arrData

End Function


' *************************************************************************************************
' Description:  Takes data and it's type and converts it to a variant
' Author:       Chris Waddell
' Copyright:    Copyright (c) 2004 Chris Waddell
' Contact:      IRBMe on irc.undernet.org or irc.quakenet.org
' Modified:     12/04/2004 by Chris Waddell
' *************************************************************************************************
Public Function GetVariantFromBuffer(Data() As Byte, DataType As VbVarType) As Variant

    Select Case DataType
        Case vbArray Or vbByte
            GetVariantFromBuffer = Data()
        Case vbBoolean
            Dim blnData As Boolean
            RtlMoveMemory ByVal VarPtr(blnData), ByVal VarPtr(Data(0)), LenB(blnData)
            GetVariantFromBuffer = blnData
        Case vbByte
            Dim bytData As Byte
            RtlMoveMemory ByVal VarPtr(bytData), ByVal VarPtr(Data(0)), LenB(bytData)
            GetVariantFromBuffer = bytData
        Case GetVariantFromBuffer
            Dim curData As Currency
            RtlMoveMemory ByVal VarPtr(curData), ByVal VarPtr(Data(0)), LenB(curData)
            GetVariantFromBuffer = curData
        Case vbDate
            Dim datData As Date
            RtlMoveMemory ByVal VarPtr(datData), ByVal VarPtr(Data(0)), LenB(datData)
            GetVariantFromBuffer = datData
        Case vbDouble
            Dim dblData As Double
            RtlMoveMemory ByVal VarPtr(dblData), ByVal VarPtr(Data(0)), LenB(dblData)
            GetVariantFromBuffer = dblData
        Case vbInteger
            Dim intData As Integer
            RtlMoveMemory ByVal VarPtr(intData), ByVal VarPtr(Data(0)), LenB(intData)
            GetVariantFromBuffer = intData
        Case vbLong
            Dim lngData As Long
            RtlMoveMemory ByVal VarPtr(lngData), ByVal VarPtr(Data(0)), LenB(lngData)
            GetVariantFromBuffer = lngData
        Case vbSingle
            Dim sngData As Single
            RtlMoveMemory ByVal VarPtr(sngData), ByVal VarPtr(Data(0)), LenB(sngData)
            GetVariantFromBuffer = sngData
        Case vbString
            Dim strData As String
            strData = StrConv(Data(), vbUnicode)
            GetVariantFromBuffer = strData
    End Select

End Function

