<%
'===== TESTING LIBRARY =====
'http://pzxc.com/asp-tdd-vbscript-test-driven-development

Dim TESTCOUNT, FAILS, T_TESTLAST, UNABRIDGED

'===== Private: Prints the testing suite CSS =====
Sub printTestingStyles
%>
 <style>
 body {font:11px Arial}
 h2,h3 {clear:left;margin:0;padding-top:10px}
 .c1 {clear:left;float:left}
 .c2 {float:left;margin-left:5px;width:40px;text-align:center}
 .c3 {float:left;margin-left:5px;}
 .c4 {float:left;margin-left:5px;}
 .pass {background-color:#00ff00}
 .warn {background-color:#ffff00}
 .fail {background-color:#ff0000}
 </style>
<%
End Sub

'===== Initializes the testing suite =====
Sub initializeTesting
	printTestingStyles
	TESTCOUNT = 0
	FAILS = 0
	T_TESTLAST = Timer()
End Sub

'===== Prints the test summary information =====
Sub summarizeTesting
%>
 <div style="clear:both;padding-top:10px">
 All tests complete. 
 There were <%=FAILS%> failures out of <%=TESTCOUNT%> tests.
 Testing took <%=timeTotal()%> seconds.
 </div>
<%
End Sub

'===== Tests that expected = actual =====
Sub assertEquals(expected, actual, message)
	Dim reason : reason = "expected " & Server.HTMLEncode(expected) & ", got " & Server.HTMLEncode(actual)
	if actual=expected then pass message,reason else fail message,reason
End Sub

'===== Tests that expected <> actual =====
sub assertNotEquals(expected, actual, message)
	Dim reason : reason = "expected !" & expected & ", got " & actual
	if actual<>expected then pass message,reason else fail message,reason
End Sub

'===== Tests that actual = true =====
sub assertTrue(actual, message)
	Dim reason : reason = "expected true, got " & actual
	if actual then pass message,reason else fail message,reason
End Sub

'===== Tests that actual = false =====
sub assertFalse(actual, message)
	Dim reason : reason = "expected false, got " & actual
	if not actual then pass message,reason else fail message,reason
End Sub

'===== Tests that actual < 20 =====
sub assertLow(actual, message)
	Dim reason : reason = "expected < 20, got " & actual
	if not (len(actual)>0 and isnumeric(actual)) or cint(actual)>=20 then
		fail message,reason
  	else
		pass message,reason
  	end if
End Sub

'===== Tests that a page has no VBScript errors and is not 404 =====
sub assertPage(relurl, message)
	Dim baseurl, contents, invalid
	baseurl = "http://" & Request.ServerVariables("SERVER_NAME")
	contents = get_url(baseurl&relurl)
	invalid = (contents="Bad URL")
	invalid = invalid or instr(1,contents,"HTTP 404")>0
	assertFalse message, invalid
End Sub

'===== Asserts a function exists and returns true if it does =====
function assertDependency(func)
	Dim r : r=isFunction(func)
	if not r then assertTrue "Dependency: "&func&"()",false
	assertDependency = r
end function

'===== Private: Returns true if the function or sub exists =====
function isFunction(func)
	on error resume next
	isFunction = IsObject(GetRef(func))
	if Err.Number<>0 then isFunction = false
	on error goto 0
end function

'===== Private: Announces a test result =====
sub announce(passfailwarn, message, reason)
	Dim thistime, which

	TESTCOUNT = TESTCOUNT + 1

	select case passfailwarn
		case 0: which="pass"
		case 1: which="warn"
		case else: which="fail" : FAILS=FAILS+1
	end select

	if not UNABRIDGED then
		if which="pass" then Response.Write "." else Response.Write "<br />"&ucase(which)&": "&message&" ("&reason&")"
	else
		thistime = timeThis(T_TESTLAST)
		Response.Write "<div class=c1>" & thistime & "</div>" & _
			"<div class=""c2 " & which & """>" & ucase(which) & "</div>" & _
			"<div class=c3>" & message & "</div>" & _
			"<div class=c4>" & reason & "</div>"
		if thistime>=5.0 then announce 1,"[Last test took more than 5 seconds.]",""
	end if
End Sub

'===== Private: Announces that a test has passed =====
sub pass(message,reason)
	announce 0, message, reason
End Sub

'===== Private: Announces that a test has failed =====
sub fail(message,reason)
	announce -1, message, reason
End Sub
%>