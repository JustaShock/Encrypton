x=msgbox("" ,0, "Encrypton .01 Alpha") 

set x = WScript.CreateObject("WScript.Shell")
mySecret = inputbox("Encrypt messager IX<OX<OS")
mySecret = StrReverse(mySecret)
x.Run "%windir%\notepad"
wscript.sleep 1000
x.sendkeys encode(mySecret)

function encode(s)
For i = 1 To Len(s)
newtxt = Mid(s, i, 1)
newtxt = Chr(Asc(newtxt)+5)
coded = coded & newtxt
Next
encode = coded
End Function