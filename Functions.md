## Serial Port VBA Functions - simplified set

##### All functions support one pre-defined Com Port only

| VBA Function                   |    Returns     | Description                                                                           |
| -------------------------------|----------------|---------------------------------------------------------------------------------------|
| `start_com_port`               | `Boolean` [^1] | Starts com port with existing settings                                                |
| `start_com_port("baud=1200")`  | `Boolean` [^1] | Starts com port with supplied settings [^5] in string                                 |
| `start_com_port(SCANNER)`      | `Boolean` [^1] | Starts com port with settings [^5] defined in VBA constant or variable SCANNER        |
| `check_com_port`               | `Long`         | Returns number of read characters waiting. -1 indicates error                         |
| `get_com_port`                 | `String`       | Returns a single waiting character string from com port                               |
| `read_com_port`                | `String`  [^3] | Returns waiting character string from com port                                        |
| `put_com_port("A")`            | `Boolean` [^1] | Send a single character string to com port                                            |
| `send_com_port("QWERTY")`      | `Boolean` [^1] | Sends [^2] supplied character string to com port                                      |
| `send_com_port(COMMANDS)`      | `Boolean` [^1] | Sends [^2] character string defined in VBA constant or variable COMMANDS to com port  |
| `send_com_port($B$5)`          | `Boolean` [^1] | Sends [^2] contents of Worksheet Cell $B$5 [^4] to com port (Excel Only)              |
| `device_ready`                 | `Boolean`      | Returns `True` if port started and Data Set Ready (DSR) input signal active           |
| `stop_com_port`                | `Boolean` [^1] | Stops com port and returns its control back to Windows                                |

##### Com Port number defined in declarations section at start of module   
`Private Const COM_PORT_NUMBER as Long = 1`    

[^1]: Function returns `True` if successful, otherwise `False`  

[^2]: Function will block until all characters are sent or write timer expires.  
      Maximum characters sent limited by timer `Write_Total_Timeout_Constant` value.   
      Long strings may cause VBA 'Not Responding' condition until transmission complete or timer expires.    
      
[^3]: Maximum characters returned = read buffer length (fixed value)    
      More waiting characters beyond buffer length may remain unread.   
      Use `check_com_port` to confirm any remaining character count if required.   
      
[^4]:  Excel will re-send if Cell $B$5 value changes     
      
[^5]: Port settings if supplied should have the same structure as the equivalent command-line Mode arguments for a COM Port
