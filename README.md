# Classic asp testing

Code from http://pzxc.com/asp-tdd-vbscript-test-driven-development

sample test

```
<!--#include virtual="/lib/unittest/aspunit.asp"-->
<!--#include virtual="/lib/utils/timer.asp"-->

<%
initializeTesting
UNABRIDGED = true

assertTrue False, "Assert False!"
assertEquals 4, 4, "4 = 4, Should not fail!"
assertEquals 4, 5, "4 != 5, Should fail!"
assertNotEquals 5, 5, "assertNotEquals(5 = 5) should fail!"

summarizeTesting
%>
```

