Public Function WRITE_COM_PORT(ByVal Write_String As String) As Long

' Important - maximum characters written may be limited by write constant timer
' Returns number of characters written

Dim Written_Length As Long
Dim Write_String_Length As Long

With COM_PORT
  
 If .Started Then
 
    Write_String_Length = Len(Write_String)

    Synchronous_Write .Handle, Write_String, Write_String_Length, Written_Length
 
 End If
  
End With

DoEvents

WRITE_COM_PORT = Written_Length

End Function
