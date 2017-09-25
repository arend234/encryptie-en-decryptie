set x = WScript.CreateObject("WScript.Shell")
     entxtde = inputbox("Enter text to be encoded")

     entxtde = StrReverse(entxtde)
   x.Run "%windir%\notepad"
     wscript.sleep 1000 
     x.sendkeys encode(entxtde)

     function encode(s) 
        For i = 1 To Len(s) 
           newtxt = Mid(s, i, 1)
           newtxt = Chr(Asc(newtxt)+3) 
           coded = coded & newtxt 
        Next 
        encode = coded 
     End Function
