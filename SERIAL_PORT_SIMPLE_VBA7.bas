Attribute VB_Name = "SERIAL_PORT_VBA_SIMPLE"
'
' https://github.com/Serialcomms/Serial-Ports-in-VBA-Simple-2022/
' https://github.com/Serialcomms/Serial-Ports-in-VBA-Simple-2022/blob/main/SERIAL_PORT_SIMPLE_VBA7.bas
'
Option Explicit
'
'--------------------------------------------------------------------------
 Private Const COM_PORT_NUMBER As Long = 1  ' < Change COM_PORT_NUMBER here
' -------------------------------------------------------------------------
'
' COM Port settings format as per command-line Mode command
' https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/mode

Private Const LONG_0  As Long = 0
Private Const LONG_1  As Long = 1
Private Const LONG_2  As Long = 2
Private Const LONG_3  As Long = 3
Private Const LONG_4  As Long = 4

Private Const HEX_0F  As Byte = &HF
Private Const HEX_20  As Byte = &H20

Private Const QUEUE_ERROR As Long = -1
Private Const HANDLE_INVALID As LongPtr = -1
Private Const READ_BUFFER_LENGTH As Long = 1024

Private Type DEVICE_CONTROL_BLOCK

             LENGTH_DCB As Long
             BAUD_RATE  As Long
             BIT_FIELD  As Long
             RESERVED   As Integer
             LIMIT_XON  As Integer
             LIMIT_XOFF As Integer
             BYTE_SIZE  As Byte
             PARITY     As Byte
             STOP_BITS  As Byte
             CHAR_XON   As Byte
             CHAR_XOFF  As Byte
             CHAR_ERROR As Byte
             CHAR_EOF   As Byte
             CHAR_EVENT As Byte
             RESERVED_1 As Integer
End Type

Private Type COM_PORT_STATUS

             BIT_FIELD As Long
             QUEUE_IN  As Long
             QUEUE_OUT As Long
End Type

Private Type COM_PORT_TIMEOUTS

             Read_Interval_Timeout          As Long
             Read_Total_Timeout_Multiplier  As Long
             Read_Total_Timeout_Constant    As Long
             Write_Total_Timeout_Multiplier As Long
             Write_Total_Timeout_Constant   As Long
End Type

Private Type COM_PORT_PROFILE

             Handle     As LongPtr
             Errors     As Long
             Signal     As Long
             Started    As Boolean
             Status     As COM_PORT_STATUS
             Timeouts   As COM_PORT_TIMEOUTS
             DCB        As DEVICE_CONTROL_BLOCK
End Type

Private COM_PORT As COM_PORT_PROFILE

Private Declare PtrSafe Function Query_Port_DCB Lib "Kernel32.dll" Alias "GetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Apply_Port_DCB Lib "Kernel32.dll" Alias "SetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Build_Port_DCB Lib "Kernel32.dll" Alias "BuildCommDCBA" (ByVal Port_Config As String, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Set_Com_Timers Lib "Kernel32.dll" Alias "SetCommTimeouts" (ByVal Port_Handle As LongPtr, ByRef Timeouts As COM_PORT_TIMEOUTS) As Boolean
Private Declare PtrSafe Function Get_Port_Modem Lib "Kernel32.dll" Alias "GetCommModemStatus" (ByVal Port_Handle As LongPtr, ByRef Modem_Status As Long) As Boolean
Private Declare PtrSafe Function Com_Port_Purge Lib "Kernel32.dll" Alias "PurgeComm" (ByVal Port_Handle As LongPtr, ByVal Port_Purge_Flags As Long) As Boolean
Private Declare PtrSafe Function Com_Port_Close Lib "Kernel32.dll" Alias "CloseHandle" (ByVal Port_Handle As LongPtr) As Boolean

Private Declare PtrSafe Function Com_Port_Clear Lib "Kernel32.dll" Alias "ClearCommError" _
(ByVal Port_Handle As LongPtr, ByRef Port_Error_Mask As Long, ByRef Port_Status As COM_PORT_STATUS) As Boolean

Private Declare PtrSafe Function Com_Port_Create Lib "Kernel32.dll" Alias "CreateFileA" _
(ByVal Port_Name As String, ByVal PORT_ACCESS As Long, ByVal SHARE_MODE As Long, ByVal SECURITY_ATTRIBUTES_NULL As Any, _
 ByVal CREATE_DISPOSITION As Long, ByVal FLAGS_AND_ATTRIBUTES As Long, Optional TEMPLATE_FILE_HANDLE_NULL) As LongPtr

Private Declare PtrSafe Function Synchronous_Read Lib "Kernel32.dll" Alias "ReadFile" _
(ByVal Port_Handle As LongPtr, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, Optional Overlapped_Null) As Boolean

Private Declare PtrSafe Function Synchronous_Write Lib "Kernel32.dll" Alias "WriteFile" _
(ByVal Port_Handle As LongPtr, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, Optional Overlapped_Null) As Boolean
'

Public Function START_COM_PORT(Optional Port_Setttings As String) As Boolean

Dim Start_Result As Boolean

With COM_PORT

 If Not .Started Then

    If OPEN_COM_PORT Then
    
        If CONFIGURE_COM_PORT(Port_Setttings) Then
            
                    .Started = True
            
                Start_Result = True
            
        Else
                STOP_COM_PORT
        End If
                       
    End If

 End If

End With

DoEvents

START_COM_PORT = Start_Result

End Function

Private Function OPEN_COM_PORT() As Boolean

Dim Temp_Open As Boolean
Dim Device_Path As String

Const OPEN_EXISTING As Long = LONG_3
Const OPEN_EXCLUSIVE As Long = LONG_0
Const SYNCHRONOUS_MODE As Long = LONG_0

Const GENERIC_RW As Long = &HC0000000
Const DEVICE_PREFIX As String = "\\.\COM"
        
Device_Path = DEVICE_PREFIX & CStr(COM_PORT_NUMBER)

COM_PORT.Handle = Com_Port_Create(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, LONG_0, OPEN_EXISTING, SYNCHRONOUS_MODE)

Temp_Open = Not (COM_PORT.Handle = HANDLE_INVALID)

OPEN_COM_PORT = Temp_Open

End Function

Private Function CONFIGURE_COM_PORT(Optional Port_Settings As String) As Boolean

Dim Temp_Result As Boolean
Dim Clean_Settings As String

Clean_Settings = CLEAN_PORT_SETTINGS(Port_Settings)

If SET_PORT_CONFIG(Clean_Settings) Then Temp_Result = SET_PORT_TIMERS
        
CONFIGURE_COM_PORT = Temp_Result

End Function

Private Function SET_PORT_CONFIG(Optional Port_Settings As String) As Boolean

Dim Temp_Result As Boolean

With COM_PORT

 If Query_Port_DCB(.Handle, .DCB) Then
  
    If Len(Port_Settings) > LONG_4 Then
    
        If Build_Port_DCB(Port_Settings, .DCB) Then
            
            Temp_Result = Apply_Port_DCB(.Handle, .DCB)
        
        End If
        
    Else
    
        Temp_Result = True
        
    End If

 End If

End With

SET_PORT_CONFIG = Temp_Result

End Function

Public Function STOP_COM_PORT() As Boolean

Dim Stop_Result As Boolean

With COM_PORT

 If .Handle > LONG_0 Then

    PURGE_COM_PORT
    
   .Started = False
    
    Stop_Result = Com_Port_Close(.Handle)
    
   .Handle = IIf(Stop_Result, LONG_0, HANDLE_INVALID)
                      
 End If

End With

DoEvents

STOP_COM_PORT = Stop_Result

End Function

Public Function READ_COM_PORT() As String

Dim Bytes_Read As Long
Dim Read_String As String
Dim Read_Buffer As String * READ_BUFFER_LENGTH  ' Important - buffer must be fixed length.

With COM_PORT
  
 If .Started Then
 
    Synchronous_Read .Handle, Read_Buffer, READ_BUFFER_LENGTH, Bytes_Read
                   
    If Bytes_Read > LONG_0 Then Read_String = Left$(Read_Buffer, Bytes_Read)
                       
 End If
  
End With

DoEvents

READ_COM_PORT = Read_String

End Function

Public Function SEND_COM_PORT(ByVal Send_String As String) As Boolean

' Important - maximum characters transmitted may be limited by write constant timer

Dim Bytes_Sent As Long
Dim Send_Result As Boolean
Dim Send_String_Length As Long

With COM_PORT
  
 If .Started Then
 
    Send_String_Length = Len(Send_String)

    Synchronous_Write .Handle, Send_String, Send_String_Length, Bytes_Sent
    
    Send_Result = (Bytes_Sent = Send_String_Length)
 
 End If
  
End With

DoEvents

SEND_COM_PORT = Send_Result

End Function

Public Function PUT_COM_PORT(ByVal Put_Character As String) As Boolean

Dim Bytes_Sent As Long
Dim Put_Result As Boolean

If COM_PORT.Started Then 

 Synchronous_Write COM_PORT.Handle, Left$(Put_Character, LONG_1), LONG_1, Bytes_Sent

End If

Put_Result = (Bytes_Sent = LONG_1)

PUT_COM_PORT = Put_Result

End Function

Public Function GET_COM_PORT() As String

Dim Bytes_Read As Long
Dim Get_String As String
Dim Read_Buffer As String * LONG_1 ' must be fixed length * 1

If COM_PORT.Started Then

 Synchronous_Read COM_PORT.Handle, Read_Buffer, LONG_1, Bytes_Read
 
 If Bytes_Read = LONG_1 Then Get_String = Read_Buffer
            
End If

GET_COM_PORT = Get_String

End Function

Public Function CHECK_COM_PORT() As Long

' Application.Volatile  ' Excel only

Dim Characters_Waiting As Long

Characters_Waiting = QUEUE_ERROR

With COM_PORT
  
 If .Started Then

    If Com_Port_Clear(.Handle, .Errors, .Status) Then Characters_Waiting = .Status.QUEUE_IN
             
 End If
             
End With
        
DoEvents

CHECK_COM_PORT = Characters_Waiting

End Function

Public Function DEVICE_READY() As Boolean

' Application.Volatile  ' Excel only

Dim Signal_State As Boolean

Const DSR_ON As Long = HEX_20

With COM_PORT
  
 If .Started Then
 
    If Get_Port_Modem(.Handle, .Signal) Then Signal_State = .Signal And DSR_ON

 End If

End With

DEVICE_READY = Signal_State

End Function

Private Function PURGE_COM_PORT() As Boolean

Dim Temp_Result As Boolean

Const PURGE_ALL As Long = HEX_0F

Temp_Result = Com_Port_Purge(COM_PORT.Handle, PURGE_ALL)

DoEvents

PURGE_COM_PORT = Temp_Result

End Function

Private Function SET_PORT_TIMERS() As Boolean

Dim Temp_Result As Boolean

Const NO_TIMEOUT As Long = -1
Const WRITE_CONSTANT As Long = 4000                           ' Should be less than approx 5000 to avoid VBA "Not Responding"
                                                              
With COM_PORT

    .Timeouts.Read_Interval_Timeout = NO_TIMEOUT              ' Timeouts not used for file reads.
    .Timeouts.Read_Total_Timeout_Constant = LONG_0            '
    .Timeouts.Read_Total_Timeout_Multiplier = LONG_0          '

    .Timeouts.Write_Total_Timeout_Constant = WRITE_CONSTANT   ' Maximum time in MilliSeconds allowed for synchronous write
    .Timeouts.Write_Total_Timeout_Multiplier = LONG_0

     Temp_Result = Set_Com_Timers(.Handle, .Timeouts)

End With

SET_PORT_TIMERS = Temp_Result

End Function

Private Function CLEAN_PORT_SETTINGS(Port_Settings As String) As String

Dim New_Settings As String

Const TEXT_COMMA As String = ","
Const TEXT_SPACE As String = " "
Const TEXT_EQUALS As String = "="
Const TEXT_DOUBLE_SPACE As String = "  "
Const TEXT_EQUALS_SPACE As String = "= "
Const TEXT_SPACE_EQUALS As String = " ="

New_Settings = Trim(Port_Settings)
New_Settings = UCase(New_Settings)

New_Settings = Replace(New_Settings, TEXT_COMMA, TEXT_SPACE, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_SPACE_EQUALS, TEXT_EQUALS, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_EQUALS_SPACE, TEXT_EQUALS, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)

CLEAN_PORT_SETTINGS = New_Settings

End Function

