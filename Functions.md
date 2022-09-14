## Serial Port VBA Functions - simplified set

##### All functions support one pre-defined Com Port only

| VBA Function                      | Description                                                                              |
| ----------------------------------|------------------------------------------------------------------------------------------|
| `start_com_port`                  | Starts[^1] com port with existing settings                                               |
| `start_com_port("baud=1200")`     | Starts[^1] com port with supplied settings[^4] in string                                 |
| `start_com_port(SCANNER)`         | Starts[^1] com port with settings[^4] defined in VBA constant or variable SCANNER        |
| `check_com_port`                  | Returns number of read characters waiting. -1 indicates error                            | 
| `put_com_port("A")`               | Sends[^1] a single character string to com port                                          |
| `get_com_port`                    | Returns a single character string from com port                                          |
| `send_com_port("QWERTY")`         | Sends[^2] supplied character string to com port                                          |
| `send_com_port(COMMANDS)`         | Sends[^2] character string defined in VBA constant or variable COMMANDS to com port      |
| `read_com_port`                   | Returns waiting character string[^3] from com port                                       |
| `device_ready`                    | Returns `True` if port started and Data Set Ready (DSR) input signal active              |
| `stop_com_port`                   | Stops[^1] com port and returns its control back to Windows                               |

##### Com Port number defined in declarations section at start of module   
`Private Const COM_PORT_NUMBER as Long = 1`    

[^1]: Function returns `True` if successful, otherwise `False`  

[^2]: Function will block until all characters are sent or write timer expires.  
      Maximum characters sent limited by timer `Write_Total_Timeout_Constant` value.   
      Long strings may cause VBA 'Not Responding' condition until transmission complete or timer expires.    
      
[^3]: Maximum characters returned = read buffer length (fixed value)    
      More waiting characters beyond buffer length may remain unread.   
      Use `check_com_port` to confirm any remaining character count if required.            
      
[^4]: Port settings if supplied should have the same structure as the equivalent command-line Mode arguments for a COM Port
