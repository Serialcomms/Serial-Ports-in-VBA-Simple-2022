Public Function CARRIER_DETECT() As Boolean

' Application.Volatile  ' Excel only

Dim Signal_State As Boolean

Const HEX_80 as Byte = &H80
Const RLSD_ON As Long = HEX_80

With COM_PORT
  
 If .Started Then
 
    If Get_Port_Modem(.Handle, .Signal) Then Signal_State = .Signal And RLSD_ON

 End If

End With

DEVICE_READY = Signal_State

End Function
  
