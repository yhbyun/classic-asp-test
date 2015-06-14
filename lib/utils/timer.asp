<%
'===== TIMING LIBRARY =====
Dim T_START : T_START = Timer()

'===== Returns the number of seconds since page began =====
Function timeTotal()
	timeTotal = formatnumber(Timer()-T_START,3)
End Function

'===== Returns the number of seconds since the given timer and resets it =====
Function timeThis(thetimer)
	Dim r : r = formatnumber(Timer()-thetimer,3)
	thetimer = Timer()
	timeThis = r
End Function

'===== Displays a timer and resets it ====
Sub showTimer(thetimer,label)
	Dim thistime
	thistime = timeThis(thetimer)
	Response.write "<div class=""timer"">" & _
		"TIMER-"&label&": "&thistime & _
		"</div>" 
End Sub
%>