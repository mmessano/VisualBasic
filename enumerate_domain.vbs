Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Windows type used to call the Net API
Private Const MAX_PREFERRED_LENGTH 	  	  As Long = -1
Private Const NERR_SUCCESS 				  As Long = 0&
Private Const ERROR_MORE_DATA 			  As Long = 234&

Private Const SV_TYPE_WORKSTATION         As Long = &H1
Private Const SV_TYPE_SERVER              As Long = &H2
Private Const SV_TYPE_SQLSERVER           As Long = &H4
Private Const SV_TYPE_DOMAIN_CTRL         As Long = &H8
Private Const SV_TYPE_DOMAIN_BAKCTRL      As Long = &H10
Private Const SV_TYPE_TIME_SOURCE         As Long = &H20
Private Const SV_TYPE_AFP                 As Long = &H40
Private Const SV_TYPE_NOVELL              As Long = &H80
Private Const SV_TYPE_DOMAIN_MEMBER       As Long = &H100
Private Const SV_TYPE_PRINTQ_SERVER       As Long = &H200
Private Const SV_TYPE_DIALIN_SERVER       As Long = &H400
Private Const SV_TYPE_XENIX_SERVER        As Long = &H800
Private Const SV_TYPE_SERVER_UNIX         As Long = SV_TYPE_XENIX_SERVER
Private Const SV_TYPE_NT                  As Long = &H1000
Private Const SV_TYPE_WFW                 As Long = &H2000
Private Const SV_TYPE_SERVER_MFPN         As Long = &H4000
Private Const SV_TYPE_SERVER_NT           As Long = &H8000
Private Const SV_TYPE_POTENTIAL_BROWSER   As Long = &H10000
Private Const SV_TYPE_BACKUP_BROWSER      As Long = &H20000
Private Const SV_TYPE_MASTER_BROWSER      As Long = &H40000
Private Const SV_TYPE_DOMAIN_MASTER       As Long = &H80000
Private Const SV_TYPE_SERVER_OSF          As Long = &H100000
Private Const SV_TYPE_SERVER_VMS          As Long = &H200000
Private Const SV_TYPE_WINDOWS             As Long = &H400000  'Windows95 and above
Private Const SV_TYPE_DFS                 As Long = &H800000  'Root of a DFS tree
Private Const SV_TYPE_CLUSTER_NT          As Long = &H1000000 'NT Cluster
Private Const SV_TYPE_TERMINALSERVER      As Long = &H2000000 'Terminal Server
Private Const SV_TYPE_DCE                 As Long = &H10000000'IBM DSS
Private Const SV_TYPE_ALTERNATE_XPORT     As Long = &H20000000'rtn alternate transport
Private Const SV_TYPE_LOCAL_LIST_ONLY     As Long = &H40000000'rtn local only
Private Const SV_TYPE_DOMAIN_ENUM         As Long = &H80000000
Private Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

Private Const SV_PLATFORM_ID_OS2       As Long = 400
Private Const SV_PLATFORM_ID_NT        As Long = 500

'Mask applied to svX_version_major in
'order to obtain the major version number.
Private Const MAJOR_VERSION_MASK        As Long = &HF

Private Type SERVER_INFO_100
  sv100_platform_id As Long
  sv100_name As Long
End Type

Private Declare Function NetServerEnum Lib "netapi32" _
  (ByVal servername As Long, _
   ByVal level As Long, _
   buf As Any, _
   ByVal prefmaxlen As Long, _
   entriesread As Long, _
   totalentries As Long, _
   ByVal servertype As Long, _
   ByVal domain As Long, _
   resume_handle As Long) As Long

Private Declare Function NetApiBufferFree Lib "netapi32" _
   (ByVal Buffer As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (pTo As Any, uFrom As Any, _
   ByVal lSize As Long)

Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long


Private Sub Form_Load()

   Command1.Caption = "Net Server Enum"

End Sub


Private Sub Command1_Click()

   Call GetServers(vbNullString)

End Sub


Private Function GetServers(sDomain As String) As Long

  'lists all servers of the specified type
  'that are visible in a domain.

   Dim bufptr          As Long
   Dim dwEntriesread   As Long
   Dim dwTotalentries  As Long
   Dim dwResumehandle  As Long
   Dim se100           As SERVER_INFO_100
   Dim success         As Long
   Dim nStructSize     As Long
   Dim cnt             As Long

   nStructSize = LenB(se100)

  'Call passing MAX_PREFERRED_LENGTH to have the
  'API allocate required memory for the return values.
  '
  'The call is enumerating all machines on the
  'network (SV_TYPE_ALL); however, by Or'ing
  'specific bit masks for defined types you can
  'customize the returned data. For example, a
  'value of 0x00000003 combines the bit masks for
  'SV_TYPE_WORKSTATION (0x00000001) and
  'SV_TYPE_SERVER (0x00000002).
  '
  'dwServerName must be Null. The level parameter
  '(100 here) specifies the data structure being
  'used (in this case a SERVER_INFO_100 structure).
  '
  'The domain member is passed as Null, indicating
  'machines on the primary domain are to be retrieved.
  'If you decide to use this member, pass
  'StrPtr("domain name"), not the string itself.
   success = NetServerEnum(0&, _
                           100, _
                           bufptr, _
                           MAX_PREFERRED_LENGTH, _
                           dwEntriesread, _
                           dwTotalentries, _
                           SV_TYPE_ALL, _
                           0&, _
                           dwResumehandle)

  'if all goes well
   If success = NERR_SUCCESS And _
      success <> ERROR_MORE_DATA Then

    'loop through the returned data, adding each
    'machine to the list
      For cnt = 0 To dwEntriesread - 1

        'get one chunk of data and cast
        'into an SERVER_INFO_100 struct
        'in order to add the name to a list
         CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize

         List1.AddItem GetPointerToByteStringW(se100.sv100_name)

      Next

   End If

  'clean up regardless of success
   Call NetApiBufferFree(bufptr)

  'return entries as sign of success
   GetServers = dwEntriesread

End Function


Public Function GetPointerToByteStringW(ByVal dwData As Long) As String

   Dim tmp() As Byte
   Dim tmplen As Long

   If dwData <> 0 Then

      tmplen = lstrlenW(dwData) * 2

      If tmplen <> 0 Then

         ReDim tmp(0 To (tmplen - 1)) As Byte
         CopyMemory tmp(0), ByVal dwData, tmplen
         GetPointerToByteStringW = tmp

     End If

   End If

End Function


