Attribute VB_Name = "TIMING_EXTRAS_VBA7"
'
Option Explicit
'
Private Declare PtrSafe Sub Kernel_Sleep_MilliSeconds Lib "Kernel32.dll" Alias "Sleep" (ByVal Sleep_MilliSeconds As Long)
Private Declare PtrSafe Sub Get_System_Time Lib "Kernel32.dll" Alias "GetSystemTime" (ByRef System_Time As VBA_SYSTEM_TIME)
Private Declare PtrSafe Function QPC Lib "Kernel32.dll" Alias "QueryPerformanceCounter" (ByRef Query_PerfCounter As Currency) As Boolean
Private Declare PtrSafe Function QPF Lib "Kernel32.dll" Alias "QueryPerformanceFrequency" (ByRef Query_Frequency As Currency) As Boolean

' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter
' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency

Private Const LONG_0 As Long = 0
Private Const LONG_1 As Long = 1
Private Const LONG_14 As Long = 14
Private Const LONG_100 As Long = 100
Private Const LONG_1000 As Long = 1000

Public Type VBA_SYSTEM_TIME

             Year           As Integer
             Month          As Integer
             WeekDay        As Integer
             Day            As Integer
             Hour           As Integer
             Minute         As Integer
             Second         As Integer
             MilliSeconds   As Integer
End Type
Public Sub KERNEL_SLEEP(Optional Wait_Milliseconds As Long)

' Sleeps for Wait_Milliseconds without causing "VBA Not Responding"

Dim Wait_Expired As Boolean
Dim Wait_Remaining As Long, Loop_Sleep_Time As Long, Effective_Wait_Time As Long

Const Loop_Time As Long = LONG_100                          ' MilliSeconds

Wait_Remaining = IIf(Wait_Milliseconds < LONG_1, LONG_1, Wait_Milliseconds)

Effective_Wait_Time = IIf(Wait_Remaining < Loop_Time, Wait_Remaining, Loop_Time)

Do

 Wait_Expired = Wait_Remaining < LONG_1
        
    If Not Wait_Expired Then

        Loop_Sleep_Time = IIf(Wait_Remaining < Effective_Wait_Time, Wait_Remaining, Effective_Wait_Time)
            
        Kernel_Sleep_MilliSeconds Loop_Sleep_Time
            
        Wait_Remaining = Wait_Remaining - Loop_Sleep_Time
            
    End If
   
 DoEvents ' prevents "VBA Not Responding"
   
Loop Until Wait_Expired

End Sub

Public Function GET_HOST_SYSTEM_TIME() As VBA_SYSTEM_TIME

' Use from VBA only.

' e.g. THIS_YEAR = GET_HOST_SYSTEM_TIME.Year

Get_System_Time GET_HOST_SYSTEM_TIME

End Function

Public Function GET_HOST_MILLISECONDS() As Long

' Application.Volatile  ' - Excel Only

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MILLISECONDS = Int(Temp_QPC)

End Function

Public Function GET_HOST_MICROSECONDS() As Currency

' Application.Volatile  ' - Excel Only

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MICROSECONDS = Int(Temp_QPC * LONG_1000)

End Function

Public Function TIMESTAMP() As String

' Application.Volatile  ' - Excel Only
' Returns VBA Time() appended with Milliseconds
' Result string extended to 14 characters.

Dim TIMESTAMP_TIME As VBA_SYSTEM_TIME

Dim TIMESTAMP_STRING As String * LONG_14

Get_System_Time TIMESTAMP_TIME

TIMESTAMP_STRING = Time() & "." & TIMESTAMP_TIME.MilliSeconds

TIMESTAMP = TIMESTAMP_STRING

End Function

