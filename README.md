<div align="center">

## Modem Comm


</div>

### Description

this code finds what comm port your modem is on. the modem must be on for it to work

you need mscomm and a textbox

thats it
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vain](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vain.md)
**Level**          |Intermediate
**User Rating**    |3.3 (13 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vain-modem-comm__1-5053/archive/master.zip)





### Source Code

```
On Error GoTo errr:
port = 1
PortinG:
MSComm1.CommPort = port
MSComm1.PortOpen = True
Form1.MSComm1.Settings = "9600,N,8,1"
MSComm1.Output = "AT" + Chr$(13)
x = 1
Do: DoEvents
x = x + 1
If x = 1000 Then MSComm1.Output = "AT" + Chr$(13)
If x = 2000 Then MSComm1.Output = "AT" + Chr$(13)
If x = 3000 Then MSComm1.Output = "AT" + Chr$(13)
If x = 4000 Then MSComm1.Output = "AT" + Chr$(13)
If x = 5000 Then MSComm1.Output = "AT" + Chr$(13)
If x = 6000 Then MSComm1.Output = "AT" + Chr$(13)
If x = 7000 Then
MSComm1.PortOpen = False
port = port + 1
GoTo PortinG:
If MSComm1.CommPort >= 5 Then
errr:
MsgBox "Can't Find Modem!"
GoTo done:
End If
End If
Loop Until MSComm1.InBufferCount >= 2
instring = MSComm1.Input
MSComm1.PortOpen = False
  Text1.Text = MSComm1.CommPort & instring
MsgBox "Modem Found On Comm" & port
done:
```

