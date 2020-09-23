Attribute VB_Name = "mdliphlpapi"
Option Explicit

'
Public Const ERROR_BUFFER_OVERFLOW = 111&
Public Const ERROR_INVALID_PARAMETER = 87
Public Const ERROR_NO_DATA = 232&
Public Const ERROR_NOT_SUPPORTED = 50&
'Public Const ERROR_SUCCESS = 0&
'
Public Const MIB_TCP_STATE_CLOSED = 1
Public Const MIB_TCP_STATE_LISTEN = 2
Public Const MIB_TCP_STATE_SYN_SENT = 3
Public Const MIB_TCP_STATE_SYN_RCVD = 4
Public Const MIB_TCP_STATE_ESTAB = 5
Public Const MIB_TCP_STATE_FIN_WAIT1 = 6
Public Const MIB_TCP_STATE_FIN_WAIT2 = 7
Public Const MIB_TCP_STATE_CLOSE_WAIT = 8
Public Const MIB_TCP_STATE_CLOSING = 9
Public Const MIB_TCP_STATE_LAST_ACK = 10
Public Const MIB_TCP_STATE_TIME_WAIT = 11
Public Const MIB_TCP_STATE_DELETE_TCB = 12
'
Private Declare Function GetIfTable Lib "IPHlpApi" (ByRef pIfRowTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function GetIpForwardTable Lib "IPHlpApi" (ByRef pIfForwardTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function SetIpForwardEntry Lib "IPHlpApi" (ByRef MIB_IPFORWARDROW As Any) As Long

 


Private Const MAX_ADAPTER_NAME_LENGTH        As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH     As Long = 8
Private Const ERROR_SUCCESS  As Long = 0
Public Const MAX_ADAPTER_NAME As Long = 128
Public Const MAX_HOSTNAME_LEN As Long = 128
Public Const MAX_DOMAIN_NAME_LEN As Long = 128
Public Const MAX_SCOPE_ID_LEN As Long = 256


Private Type IP_ADDRESS_STRING
    ipaddr(0 To 15)  As Byte
End Type

Private Type IP_MASK_STRING
    IpMask(0 To 15)  As Byte
End Type


Private Type MIB_IPFORWARDROW
  dwForwardDest(0 To 3)         As Byte ' IP addr of destination
   dwForwardMask(0 To 3)          As Byte ' subnetwork mask of destination
   dwForwardPolicy       As Long ' conditions for multi-path route
   dwForwardNextHop(0 To 3)               As Byte ' IP address of next hop = passerel
   dwForwardIfIndex      As Long ' index of interface
   dwForwardType         As Long ' route type
   dwForwardProto        As Long ' protocol that generated route
   dwForwardAge          As Long ' age of route
   dwForwardNextHopAS    As Long ' autonomous system number of next hop
   dwForwardMetric1      As Long ' protocol-specific metric
   dwForwardMetric2      As Long ' protocol-specific metric
   dwForwardMetric3      As Long ' protocol-specific metric
   dwForwardMetric4      As Long ' protocol-specific metric
   dwForwardMetric5      As Long ' protocol-specific metric

End Type

Private Type MIB_IPFORWARDTABLE
  dwNumEntries As Long    ' number of entries in the table
  MIB_IPFORWARDROW    As MIB_IPFORWARDROW  ' array of route entries

End Type
'Author : Erwan L.
'mail:erwan.l@free.fr
Public Sub routeprint(ByRef infos() As String)
Dim lngRetVal As Long
Dim arrBuffer()     As Byte
Dim lngSize         As Long
Dim lngRows As Long
lngRetVal = GetIpForwardTable(ByVal 0&, lngSize, 0&)
If lngRetVal = ERROR_NOT_SUPPORTED Then
        'This API works only on Win 98/2000 and NT4 with SP4
        MsgBox "IP Helper is not supported by this system."
        Exit Sub
End If
'Prepare the buffer
ReDim arrBuffer(0 To lngSize - 1) As Byte
'
'And call the function one more time
lngRetVal = GetIpForwardTable(arrBuffer(0), lngSize, 0)
'
If lngRetVal = ERROR_SUCCESS Then
Dim IPFORWARDTABLE As MIB_IPFORWARDTABLE     '
Dim ipforwardrow As MIB_IPFORWARDROW
Dim i As Integer
Dim str As String
'The first 4 bytes (the Long value) contain the quantity of the table rows
CopyMemory_any lngRows, arrBuffer(0), 4
ReDim infos(lngRows)
For i = 1 To lngRows
    CopyMemory_any ipforwardrow, arrBuffer(4 + (i - 1) * Len(ipforwardrow)), Len(ipforwardrow)
    infos(i) = _
    ipforwardrow.dwForwardDest(0) & "." & ipforwardrow.dwForwardDest(1) & "." & ipforwardrow.dwForwardDest(2) & "." & ipforwardrow.dwForwardDest(3) & _
    vbTab & ipforwardrow.dwForwardMask(0) & "." & ipforwardrow.dwForwardMask(1) & "." & ipforwardrow.dwForwardMask(2) & "." & ipforwardrow.dwForwardMask(3) & _
    vbTab & ipforwardrow.dwForwardNextHop(0) & "." & ipforwardrow.dwForwardNextHop(1) & "." & ipforwardrow.dwForwardNextHop(2) & "." & ipforwardrow.dwForwardNextHop(3) & _
    vbTab & ipforwardrow.dwForwardIfIndex
   
Next i
End If
End Sub

Public Function TrimNull2(item As String)
    TrimNull2 = LTrim(RTrim(Replace(item, Chr(0), "")))
End Function
