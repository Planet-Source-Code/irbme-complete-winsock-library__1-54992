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
' A list of error codes which may be returned by various winsock functions.
' *************************************************************************************************
Public Enum ErrorCodeType

    ' The operation completed successfuly
    NoError = 0

    ' A blocking operation was interrupted by a call to WSACancelBlockingCall.
    InterruptedFunctionCall = WSAEINTR

    ' An attempt was made to access a socket in a way forbidden by its access permissions.
    ' An example is using a broadcast address for sendto without broadcast permission being
    ' set using setsockopt(SO_BROADCAST).
    '
    ' Another possible reason for the WSAEACCES error is that when the bind function is called
    ' (on Windows NT 4 SP4 or later), another application, service, or kernel mode driver is
    ' bound to the same address with exclusive access. Such exclusive access is a new feature
    ' of Windows NT 4 SP4 and later, and is implemented by using the SO_EXCLUSIVEADDRUSE option.
    PermissionDenied = WSAEACCES

    ' The system detected an invalid pointer address in attempting to use a pointer argument of
    ' a call. This error occurs if an application passes an invalid pointer value, or if the
    ' length of the buffer is too small. For instance, if the length of an argument, which is a
    ' sockaddr structure, is smaller than the sizeof(sockaddr).
    BadAddress = WSAEFAULT
    
    ' Some invalid argument was supplied (for example, specifying an invalid level to the
    ' setsockopt function). In some instances, it also refers to the current state of the
    ' socket—for instance, calling accept on a socket that is not listening.
    InvalidArgument = WSAEINVAL
    
    ' Each implementation may have a maximum number of socket handles available, either globally,
    ' per process, or per thread.
    TooManyOpenFiles = WSAEMFILE

    ' This error is returned from operations on nonblocking sockets that cannot be completed
    ' immediately, for example recv when no data is queued to be read from the socket. It is a
    ' nonfatal error, and the operation should be retried later. It is normal for WSAEWOULDBLOCK
    ' to be reported as the result from calling connect on a nonblocking SOCK_STREAM socket,
    ' since some time must elapse for the connection to be established.
    ResourceTemporarilyUnavailable = 10035

    ' A blocking operation is currently executing. Windows Sockets only allows a single blocking
    ' operation—per- task or thread—to be outstanding, and if any other function call is made
    ' (whether or not it references that or any other socket) the function fails with the
    ' WSAEINPROGRESS error.
    OperationNowInProgress = WSAEINPROGRESS
    
    ' An operation was attempted on a nonblocking socket with an operation already in
    ' progress—that is, calling connect a second time on a nonblocking socket that is already
    ' connecting, or canceling an asynchronous request (WSAAsyncGetXbyY) that has already been
    ' canceled or completed.
    OperationAlreadyInProgress = WSAEALREADY
    
    ' An operation was attempted on something that is not a socket. Either the socket handle
    ' parameter did not reference a valid socket, or for select, a member of an fd_set was not
    ' valid.
    SocketOperationOnNonSocket = WSAENOTSOCK
    
    ' A required address was omitted from an operation on a socket. For example, this error is
    ' returned if sendto is called with the remote address of ADDR_ANY.
    DestinationAddressRequired = WSAEDESTADDRREQ

    ' A message sent on a datagram socket was larger than the internal message buffer or some
    ' other network limit, or the buffer used to receive a datagram was smaller than the
    ' datagram itself.
    MessageTooLong = WSAEMSGSIZE

    ' A protocol was specified in the socket function call that does not support the semantics
    ' of the socket type requested. For example, the ARPA Internet UDP protocol cannot be
    ' specified with a socket type of SOCK_STREAM.
    ProtocolWrongTypeForSocket = 10041
    
    ' An unknown, invalid or unsupported option or level was specified in a getsockopt or
    ' setsockopt call.
    BadProtocolOption = WSAENOPROTOOPT
    
    ' The requested protocol has not been configured into the system, or no implementation for
    ' it exists. For example, a socket call requests a SOCK_DGRAM socket, but specifies a stream
    ' protocol.
    ProtocolNotSupported = WSAEPROTONOSUPPORT

    ' The support for the specified socket type does not exist in this address family.
    ' For example, the optional type SOCK_RAW might be selected in a socket call, and the
    ' implementation does not support SOCK_RAW sockets at all.
    SocketTypeNotSupported = WSAESOCKTNOSUPPORT

    ' The attempted operation is not supported for the type of object referenced. Usually this
    ' occurs when a socket descriptor to a socket that cannot support this operation is trying
    ' to accept a connection on a datagram socket.
    OperationNotSupported = WSAEOPNOTSUPP

    ' The protocol family has not been configured into the system or no implementation for it
    ' exists. This message has a slightly different meaning from WSAEAFNOSUPPORT. However, it
    ' is interchangeable in most cases, and all Windows Sockets functions that return one of
    ' these messages also specify WSAEAFNOSUPPORT.
    ProtocolFamilyNotSupported = WSAEPFNOSUPPORT

    ' An address incompatible with the requested protocol was used. All sockets are created with
    ' an associated address family (that is, AF_INET for Internet Protocols) and a generic
    ' protocol type (that is, SOCK_STREAM). This error is returned if an incorrect protocol is
    ' explicitly requested in the socket call, or if an address of the wrong family is used for
    ' a socket, for example, in sendto.
    AddressFamilyNotSupportedByProtocolFamily = WSAEAFNOSUPPORT

    ' Typically, only one usage of each socket address (protocol/IP address/port) is permitted.
    ' This error occurs if an application attempts to bind a socket to an IP address/port that
    ' has already been used for an existing socket, or a socket that was not closed properly, or
    ' one that is still in the process of closing. For server applications that need to bind
    ' multiple sockets to the same port number, consider using setsockopt (SO_REUSEADDR).
    ' Client applications usually need not call bind at all—connect chooses an unused port
    ' automatically. When bind is called with a wildcard address (involving ADDR_ANY), a
    ' WSAEADDRINUSE error could be delayed until the specific address is committed. This could
    ' happen with a call to another function later, including connect, listen, WSAConnect, or
    ' WSAJoinLeaf.
    AddressAlreadyInUse = WSAEADDRINUSE

    ' The requested address is not valid in its context. This normally results from an attempt to
    ' bind to an address that is not valid for the local computer. This can also result from
    ' connect, sendto, WSAConnect, WSAJoinLeaf, or WSASendTo when the remote address or port is
    ' not valid for a remote computer (for example, address or port 0).
    CannotAssignRequestedAddress = WSAEADDRNOTAVAIL

    ' A socket operation encountered a dead network. This could indicate a serious failure of the
    ' network system (that is, the protocol stack that the Windows Sockets DLL runs over), the
    ' network interface, or the local network itself.
    NetworkIsDown = WSAENETDOWN

    ' A socket operation was attempted to an unreachable network. This usually means the local
    ' software knows no route to reach the remote host.
    NetworkIsUnreachable = WSAENETUNREACH
    
    ' The connection has been broken due to keep-alive activity detecting a failure while the
    ' operation was in progress. It can also be returned by setsockopt if an attempt is made to
    ' set SO_KEEPALIVE on a connection that has already failed.
    NetworkDroppedConnectionOnReset = WSAENETRESET

    ' An established connection was aborted by the software in your host computer, possibly due
    ' to a data transmission time-out or protocol error.
    SoftwareCausedConnectionAbort = WSAECONNABORTED

    ' An existing connection was forcibly closed by the remote host. This normally results if the
    ' peer application on the remote host is suddenly stopped, the host is rebooted, the host or
    ' remote network interface is disabled, or the remote host uses a hard close (see setsockopt
    ' for more information on the SO_LINGER option on the remote socket). This error may also
    ' result if a connection was broken due to keep-alive activity detecting a failure while one
    ' or more operations are in progress. Operations that were in progress fail with WSAENETRESET.
    ' Subsequent operations fail with WSAECONNRESET.
    ConnectionResetByPeer = WSAECONNRESET

    ' An operation on a socket could not be performed because the system lacked sufficient buffer
    ' space or because a queue was full.
    NoBufferSpaceAvailable = WSAENOBUFS

    ' A connect request was made on an already-connected socket. Some implementations also
    ' return this error if sendto is called on a connected SOCK_DGRAM socket (for SOCK_STREAM
    ' sockets, the to parameter in sendto is ignored) although other implementations treat this
    ' as a legal occurrence.
    SocketIsAlreadyConnected = WSAEISCONN

    ' A request to send or receive data was disallowed because the socket is not connected and
    ' (when sending on a datagram socket using sendto) no address was supplied. Any other type
    ' of operation might also return this error—for example, setsockopt setting SO_KEEPALIVE if
    ' the connection has been reset.
    SocketIsNotConnected = WSAENOTCONN

    ' A request to send or receive data was disallowed because the socket had already been shut
    ' down in that direction with a previous shutdown call. By calling shutdown a partial close
    ' of a socket is requested, which is a signal that sending or receiving, or both have been
    ' discontinued.
    CannotSendAfterSocketShutdown = WSAESHUTDOWN

    ' A connection attempt failed because the connected party did not properly respond after a
    ' period of time, or the established connection failed because the connected host has failed
    ' to respond.
    ConnectionTimedOut = WSAETIMEDOUT

    ' No connection could be made because the target computer actively refused it. This usually
    ' results from trying to connect to a service that is inactive on the foreign host—that is,
    ' one with no server application running.
    ConnectionRefused = WSAECONNREFUSED

    ' A socket operation failed because the destination host is down. A socket operation
    ' encountered a dead host. Networking activity on the local host has not been initiated.
    ' These conditions are more likely to be indicated by
    HostIsDown = WSAEHOSTDOWN

    ' A socket operation was attempted to an unreachable host. See WSAENETUNREACH.
    NoRouteToHost = WSAEHOSTUNREACH

    ' A Windows Sockets implementation may have a limit on the number of applications that can
    ' use it simultaneously. WSAStartup may fail with this error if the limit has been reached.
    TooManyProcesses = WSAEPROCLIM

    ' This error is returned by WSAStartup if the Windows Sockets implementation cannot function
    ' at this time because the underlying system it uses to provide network services is currently
    ' unavailable. Users should check:
    '
    ' - That the appropriate Windows Sockets DLL file is in the current path.
    '
    ' - That they are not trying to use more than one Windows Sockets implementation
    '   simultaneously. If there is more than one Winsock DLL on your system, be sure the first
    '   one in the path is appropriate for the network subsystem currently loaded.
    '
    ' - The Windows Sockets implementation documentation to be sure all necessary components are
    '   currently installed and configured correctly.
    '
    NetworkSubsystemIsUnavailable = WSASYSNOTREADY

    ' The current Windows Sockets implementation does not support the Windows Sockets
    ' specification version requested by the application. Check that no old Windows Sockets DLL
    ' files are being accessed.
    WinsockDllVersionOutOfRange = WSAVERNOTSUPPORTED

    ' Either the application has not called WSAStartup or WSAStartup failed. The application may
    ' be accessing a socket that the current active task does not own (that is, trying to share a
    ' socket between tasks), or WSACleanup has been called too many times.
    SuccessfulWSAStartupNotYetPerformed = WSANOTINITIALISED

    ' Returned by WSARecv and WSARecvFrom to indicate that the remote party has initiated a
    ' graceful shutdown sequence.
    GracefulShutdownInProgress = WSAEDISCON

    ' The specified class was not found.
    ClassTypeNotFound = WSATYPE_NOT_FOUND

    ' No such host is known. The name is not an official host name or alias, or it cannot be
    ' found in the database(s) being queried. This error may also be returned for protocol and
    ' service queries, and means that the specified name could not be found in the relevant
    ' database.
    HostNotFound = WSAHOST_NOT_FOUND

    ' This is usually a temporary error during host name resolution and means that the local
    ' server did not receive a response from an authoritative server. A retry at some time later
    ' may be successful.
    NonAuthoritativeHostNotFound = WSATRY_AGAIN

    ' This indicates some sort of nonrecoverable error occurred during a database lookup. This
    ' may be because the database files (for example, BSD-compatible HOSTS, SERVICES, or
    ' PROTOCOLS files) could not be found, or a DNS request was returned by the server with a
    ' severe error.
    NonRecoverableError = WSANO_RECOVERY

    ' The requested name is valid and was found in the database, but it does not have the
    ' correct associated data being resolved for. The usual example for this is a host
    ' name-to-address translation attempt (using gethostbyname or WSAAsyncGetHostByName)
    ' which uses the DNS (Domain Name Server). An MX record is returned but no A
    ' record—indicating the host itself exists, but is not directly reachable.
    ValidNameButNoDataRecordOfRequestedType = WSANO_DATA

    ' Specified event object handle is invalid.
    ' An application attempts to use an event object, but the specified handle is not valid.
    InvalidHandle = WSA_INVALID_HANDLE ' OS Dependant

    ' An application used a Windows Sockets function which directly maps to a Windows function.
    ' The Windows function is indicating a problem with one or more parameters.
    InvalidParameter = WSA_INVALID_PARAMETER ' OS Dependant

    ' Overlapped I/O event object not in signaled state.
    ' The application has tried to determine the status of an overlapped operation which is not
    ' yet completed. Applications that use WSAGetOverlappedResult (with the fWait flag set to
    ' FALSE) in a polling mode to determine when an overlapped operation has completed.
    IOEventIncomplete = WSA_IO_INCOMPLETE  ' OS Dependant

    ' Overlapped operations will complete later.
    ' The application has initiated an overlapped operation that cannot be completed immediately.
    ' A completion indication will be given later when the operation has been completed.
    IOOperationPending = WSA_IO_PENDING ' OS Dependant

    ' Insufficient memory available.
    ' An application used a Windows Sockets function that directly maps to a Windows function.
    ' The Windows function is indicating a lack of required memory resources.
    InsufficientMemory = WSA_NOT_ENOUGH_MEMORY ' OS Dependant

    ' Overlapped operation aborted.
    ' An overlapped operation was canceled due to the closure of the socket, or the execution of
    ' the SIO_FLUSH command in WSAIoctl.
    OverlappedOperationAborted = WSA_OPERATION_ABORTED ' OS Dependant

    ' Invalid procedure table from service provider.
    ' A service provider returned a bogus procedure table to Ws2_32.dll. (Usually caused by one
    ' or more of the function pointers being null.)
    InvalidProcedureTable = WSAEINVALIDPROCTABLE ' OS Dependant

    ' Invalid service provider version number.
    ' A service provider returned a version number other than 2.0.
    InvalidServiceProviderVersion = WSAEINVALIDPROVIDER ' OS Dependant

    ' Unable to initialize a service provider.
    ' Either a service provider's DLL could not be loaded (LoadLibrary failed) or the provider's
    ' WSPStartup/NSPStartup function failed.
    UnableToInitializeServiceProvider = WSAEPROVIDERFAILEDINIT ' OS Dependant

    ' System call failure.
    ' Generic error code, returned under various conditions.
    '
    ' - Returned when a system call that should never fail does fail. For example, if a call to
    '   WaitForMultipleEvents fails or one of the registry functions fails trying to manipulate
    '   the protocol/namespace catalogs.
    '
    ' - Returned when a provider does not return SUCCESS and does not provide an extended error
    '   code. Can indicate a service provider implementation error.
    SystemCallFailure = WSASYSCALLFAILURE ' OS Dependant
End Enum


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
