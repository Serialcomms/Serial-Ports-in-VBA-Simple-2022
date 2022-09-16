# Serial Port functions in VBA Simple 2022

This is the simplified set of VBA Serial Port functions for use with **one pre-defined Com Port only**.

Intended for use with serial devices which have a well-defined set of short commands and responses. 

No debug or extended transmit & receive functionality is provided. 

No other files, licences, payments or registrations are required.  

Functions can be used directly in Excel cells where appropriate. 


<details>

<summary>Com Ports</summary>

<p>
  
- Functions work with both Hardware and Virtual (software) Com Port types 
 
- All API functions are `'Synchronous'` as some port types do not respond correctly in `'Overlapped'` mode  

</p>

</details>

<details>

<summary>Read Functions</summary>

<p>
  
_Assume that all data has already been sent by the attached serial device and is ready waiting to be read_

- `check_com_port` can be used to confirm expected number of characters are waiting before committing read 

- No pre or post read delays for any in-flight data reception to complete are provided.
  
- Data will be read in one synchronous API call.
  
- Maximum characters per read call = `READ_BUFFER_LENGTH`
  
- `check_com_port` function can be used again to check for any new or remaining characters. 
    
</p>

</details>

<details>
  
<summary>Write Functions</summary>
 
<p>

Writes are synchronous and functions can block until outgoing data is processed or write timer expires 
    
- Short strings will return quickly as data is buffered for transmission    
- Maximum number of characters sent is limited by write timer value in milliseconds
- Character limit per send is approximately = ( Baud Rate * WRITE_CONSTANT ) / 10000

</p>

</details>

<details>
  
<summary>Control Functions</summary>

<p>

### Com Port Start, Stop ###
  
- Allow a few MilliSeconds for functions to return and for any attached hardware to stabilise   
- Functions return `True` or `False` to indicate success or failure  
  
### Data Waiting Check ###
  
- Function returns number of characters waiting to be read   
- Return number can be zero if no data waiting  
- Return value of -1 indicates error, including port not started 
  
### Device Ready Check ###  

- Function returns `True` if port started and **Data Set Ready** input signal is active 
   
</p>  
  
</details>

[Function List Table](Functions.md)
