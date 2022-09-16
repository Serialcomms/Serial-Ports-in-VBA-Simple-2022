Public Function CLEAR_TO_SEND() As Boolean

' Application.Volatile  ' Excel only

Dim Signal_State As Boolean

Const HEX_10 as Byte = &H10
Const CTS_ON As Long = HEX_10

With COM_PORT
  
 If .Started Then
 
    If Get_Port_Modem(.Handle, .Signal) Then Signal_State = .Signal And CTS_ON

 End If

End With

CLEAR_TO_SEND = Signal_State

End Function
