x=msgbox("" ,0, "{Decrypton .01 Alpha} is coressponding version") 

set x = WScript.CreateObject("WScript.Shell")
mySecret = inputbox("Decryptor for Encrypton 0.1 Alpha")
mySecret = StrReverse(mySecret)
x.Run "%windir%\notepad"
wscript.sleep 1000
x.sendkeys encode(mySecret)

function encode(s)
For i = 1 To Len(s)
newtxt = Mid(s, i, 1)
newtxt = Chr(Asc(newtxt)-5)
coded = coded & newtxt
Next
encode = coded
End Function