Public Function DEVICE_CALLING() As Boolean

' Application.Volatile  ' Excel only

Dim Signal_State As Boolean

Const HEX_40 as Byte = &H40
Const RING_ON As Long = HEX_40

With COM_PORT
  
 If .Started Then
 
    If Get_Port_Modem(.Handle, .Signal) Then Signal_State = .Signal And RING_ON

 End If

End With

DEVICE_CALLING = Signal_State

End Function
