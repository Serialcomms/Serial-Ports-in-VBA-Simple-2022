Public Function WRITE_COM_PORT(ByVal Write_String As String) As Long

' Important - maximum characters written may be limited by write constant timer
' Returns number of characters written

Dim Bytes_Sent As Long
Dim Write_String_Length As Long

With COM_PORT
  
 If .Started Then
 
    Write_String_Length = Len(Write_String)

    Synchronous_Write .Handle, Write_String, Write_String_Length, Bytes_Sent
 
 End If
  
End With

DoEvents

WRITE_COM_PORT = Bytes_Sent

End Function
