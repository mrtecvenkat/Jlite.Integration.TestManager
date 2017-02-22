on error resume next
'msgbox "wait0"
Set args = WScript.Arguments
Set WshShell = WScript.CreateObject("WScript.Shell")
iaction = args.Item(0)  
'iaction  = inputbox("please enter type")
'ivalue = inputbox("Please enter the value")
'ivalue2 = inputbox("Please enter the value")
ivalue = args.Item(1)  
ivalue2 = args.Item(2)
If lcase(trim(iaction)) = "browse" Then
call doHandleBrowse(ivalue)
Else if lcase(trim(iaction)) = "dialog" Then
call doHandleDialog(ivalue)
Else if lcase(trim(iaction)) = "submit" Then
call doHandleSubmit(ivalue)
Else if lcase(trim(iaction)) = "selectoffer" Then
call doHandleOfferSelect()
Else if lcase(trim(iaction)) = "selectandsubmit" Then
call doHandleSubmitWithSelect(ivalue,ivalue2)
Else if lcase(trim(iaction)) = "selectpricing" Then
call doHandlePriceSelect()
Else if lcase(trim(iaction)) = "browse1" Then
call doHandleBrowse1(ivalue)
Else if lcase(trim(iaction)) = "selectandsubmitlob" Then
call doHandleSubmitWithSelectlob(ivalue,ivalue2)
Else if lcase(trim(iaction)) = "selectandsubmit1" Then
call doHandleSubmitWithSelecttwotimes(ivalue,ivalue2)
Else if lcase(trim(iaction)) = "selectandsubmitacq" Then
call doHandleSubmitWithSelectACQ(ivalue,ivalue2)
End if
End If
End If
End If
End If
End If
End If
End If
End If
End If

Function doHandleOfferSelect()
	WScript.Sleep 10000
	WshShell.SendKeys " "
    WScript.Sleep 10000
	WshShell.SendKeys " "
	'msgbox "dpme"
End Function
Function doHandlePriceSelect()
	WScript.Sleep 10000
'	WshShell.SendKeys "{TAB}"
'        WScript.Sleep 2000
	WshShell.SendKeys " "       
        WScript.Sleep 10000
	WshShell.SendKeys " "       

End Function

Function doHandleBrowse(path)
	WScript.Sleep 10000
	WshShell.SendKeys path 
    WScript.Sleep 2000
    WshShell.SendKeys "{TAB}{TAB}"
	WScript.Sleep 2000
    WshShell.SendKeys "{ENTER}"
    
End Function

Function doHandleDialog(details)

WScript.Sleep 15000

If instr(details,"all")>0 Then
WshShell.SendKeys "{TAB}{TAB}{TAB}"
'WScript.Sleep 2000
WshShell.SendKeys "{ENTER}"
'WScript.Sleep 5000          
WshShell.SendKeys "{TAB}{TAB}{TAB}"
'WScript.Sleep 2000 
WshShell.SendKeys "{ENTER}"
'WScript.Sleep 2000 
'WshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
'WshShell.SendKeys "{ENTER}"

Else If instr(details,",")>0 Then

allOptions = split(details,",")

for i=0 to ubound(allOptions)
 if i=0 then
	WshShell.SendKeys allOptions(i)
	WshShell.SendKeys "{TAB}"
	WScript.Sleep 2000
	WshShell.SendKeys "{ENTER}"
 else	
	WshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
	WshShell.SendKeys allOptions(i)
	WshShell.SendKeys "{TAB}"
	WScript.Sleep 2000                      

	WshShell.SendKeys "{ENTER}"
	

 end if
	
Next
	WshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}"
	WScript.Sleep 2000
	WshShell.SendKeys "{ENTER}"


Else
	WshShell.SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
	WScript.Sleep 2000
	WshShell.SendKeys "{ENTER}" 
End If
End If

End Function

Function doHandleSubmit(comments)
WScript.Sleep 15000
WshShell.SendKeys comments
WScript.Sleep 2000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 15000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"
End Function   

Function doHandleBrowse1(path)
'    msgbox "wait0" & path
    WScript.Sleep 10000
   WshShell.SendKeys "{TAB}{TAB}"
    WshShell.SendKeys " "
	WScript.Sleep 10000
	WshShell.SendKeys path 
    WScript.Sleep 2000
    WshShell.SendKeys "{TAB}{TAB}"
	WScript.Sleep 2000
    WshShell.SendKeys "{ENTER}"
      WScript.Sleep 2000
    WshShell.SendKeys "{TAB}{TAB}{TAB}"
	WScript.Sleep 2000
    WshShell.SendKeys "{ENTER}"    
End Function
 
Function doHandleSubmitWithSelect(ioption,icomments)
WScript.Sleep 15000
'WshShell.SendKeys "{TAB}"
WshShell.SendKeys ioption
WScript.Sleep 5000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
'WScript.Sleep 2000
WshShell.SendKeys icomments
'WScript.Sleep 2000
WshShell.SendKeys "{TAB}"
WScript.Sleep 2000
WshShell.SendKeys "{ENTER}"
End Function 

Function doHandleSubmitWithSelectlob(ioption,icomments)
WScript.Sleep 15000
'WshShell.SendKeys "{TAB}"
WshShell.SendKeys ioption
WScript.Sleep 5000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
'WScript.Sleep 2000
WshShell.SendKeys icomments
'WScript.Sleep 2000
WshShell.SendKeys "{TAB}"
WScript.Sleep 2000
WshShell.SendKeys "{ENTER}"
'Script.Sleep 2000
'shShell.SendKeys "{TAB}{TAB}{TAB}"
'Script.Sleep 2000
'WshShell.SendKeys "{ENTER}"
'WScript.Sleep 2000
'WshShell.SendKeys "{ENTER}"
End Function 

Function doHandleSubmitWithSelecttwotimes(ioption,icomments)
WScript.Sleep 15000
'WshShell.SendKeys "{TAB}"
WshShell.SendKeys ioption
WScript.Sleep 5000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
'WScript.Sleep 2000
WshShell.SendKeys icomments
'WScript.Sleep 2000
WshShell.SendKeys "{TAB}"
WScript.Sleep 2000
WshShell.SendKeys "{ENTER}"
'WScript.Sleep 15000
'WshShell.SendKeys "{TAB}"
'WshShell.SendKeys "{TAB}"
'WshShell.SendKeys "{TAB}"
'WScript.Sleep 2000
'WshShell.SendKeys "{ENTER}"
'WScript.Sleep 2000
'WshShell.SendKeys "{ENTER}"
End Function 

Function doHandleSubmitWithSelectACQ(ioption,icomments)
WScript.Sleep 15000
WshShell.SendKeys ioption
WScript.Sleep 5000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys icomments
WshShell.SendKeys "{TAB}"
WScript.Sleep 2000
WshShell.SendKeys "{ENTER}"
'WScript.Sleep 15000
'WshShell.SendKeys "%"
'WshShell.SendKeys "{F4}"
End Function

Set WshShell = nothing