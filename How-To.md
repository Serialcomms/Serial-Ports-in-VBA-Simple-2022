## Installing and Testing Simple VBA Module

<details>

<summary>Before Starting</summary>  

<p>
  
- Before starting, note that these functions are intended for use on Windows PC only. 
    
- Required Microsoft Office applications (Excel, Word) should be installed on the PC. 
    
- Functions will not work with Online Office versions such as Office 365, Outlook etc.   

- Macros must be enabled and documents should be saved as type Macro-Enabled. 

- An 'intermediate' or better level of VBA and serial comms knowledge is assumed. 

- Host PC should have a physical or virtual Com Port available for use.  
    
- Testing requires a second device to be connected to the Com Port. 
  
- Testing assumes that all hardware, cabling and devices are configured and working correctly.   
  
</details>

</p>  

<details>
    
<summary>Versions</summary>   
   
<p>
  
There are two versions of the file available 
  
  1. `SERIAL_PORT_SIMPLE_VBA6.bas` - for use with pre Office 2010 editions
  2. `SERIAL_PORT_SIMPLE_VBA7.bas` - for use with Office 2010+ editions
  
     _VBA7 version uses `PtrSafe` function definitions and `LongPtr` variables_
  
</p>
  
</details>
  
<details>
    
<summary>Installing</summary> 
  
<p>  

- Download SERIAL_PORT_SIMPLE_VBAn.bas to a known location on your PC  
- Open a new Office Application (Excel,Word,Access) document   
- Enter the  Office Application's VBA Environment (Alt-F11)  
- From VBA Environment, view the Project Explorer (Control-R)  
- From Project Explorer, right-hand click and select Import File  
- Import the file SERIAL_PORT_SIMPLE_VBAn.bas  
- Check that a new module SERIAL_PORT_VBA_SIMPLE is created and visible in the Modules folder
- Check/edit constant `COM_PORT_NUMBER` value in module SERIAL_PORT_VBA_SIMPLE.  
- Close and return to Office application (Alt-Q)  
- IMPORTANT - save document as type Macro-Enabled with a file name of your choice  
    
 </details>
 
 </p>  
   
 <details>
  
 <summary>Initial Testing</summary>   
 
 <p>  
 
 To test, another PC or serial device should be connected to the host PC's Com Port.  
     
 Reconfirm `COM_PORT_NUMBER` value in module declarations section is correct.    
    
- Re-enter the VBA Environment (Alt-F11)  
- Select the VBA Immediate Window (Control-G)  
- Enter the following commands in the Immediate Window :-  
-  `?start_com_port` to use default Com Port settings  - or -  
-  `?start_com_port("Baud=9600 Data=8")` to specify Com Port settings  
- Check that `True` is returned by the function in the Immediate window
- Send some text to your device from VBA e.g. 
-  `?send_com_port("Hello Device")`
- Check that 'Hello Device' is received correctly on your device
- On your device, send some text back to VBA - e.g. `Hello Excel`  
- Read the text from your device back to VBA
-  `?read_com_port`
- Check that `Hello Excel` is displayed correctly in the Immediate Window
   
 </details>
 
 </p>
 
 <details>
 
 <summary>Optional Testing</summary>     
 
 <p>
   
- Assign some text to variable Y in the Immediate Window 
-  `Y = "Excel to Device test on " & date`  
- Send variable Y to your device and check that it is received correctly
- `?send_com_port(Y)`
- On your device, send some more text back to VBA - e.g. `QWERTY` 
- Receive device text into variable Z
-  `Z = read_com_port`
- Display variable Z and check if text received correctly
-  `?Z`
- Optional - explore other Public functions, e.g. 
-  `?check_com_port`
-  `?put_com_port("A")`
-  `?get_com_port`
-  `?device_ready`
- Close the com port
- `?stop_com_port`
   
 </details> 
 </p>  

<details>
 
<summary>Summary</summary>

<p>    
    
- Public functions can be incorporated into your own VBA programs   
- Private functions are not intended to be called by end VBA users    
- See other serialcomms repositories for Ribbon customisation information   
- Where appropriate, Public functions can also be used directly in Excel worksheet Cells  

</details>   
</p>
