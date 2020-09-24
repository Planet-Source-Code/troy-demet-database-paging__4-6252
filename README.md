<div align="center">

## Database Paging


</div>

### Description

I needed to be able to page through a database of about 1000 records and I was unable to use .AbsolutePage so I had to try something else.
 
### More Info
 
You will need to make a connection to your own database and re-name to records.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy Demet](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-demet.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-demet-database-paging__4-6252/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<!--#include File="_fpclass/adovbs.inc"-->
<%
	gend = CStr(Request.QueryString("gender"))
' **************** Added July 5, 2000 ************
Dim iPageSize 'How many records to show
Dim iRecCurrent ' The page we want to show
Dim sSQL   'SQL command to execute
Dim RecSet  	'The ADODB recordset object
Dim I   'Standard looping var
Dim iRecEnd	' Last Record
Dim iRecMax	' Max of record loop
Dim J		' Loop variabel
Dim iRecNext	' Var of Next record to start at
Dim iRecPrev	' Var of Previous record
Dim sGender	' Var for displaying whether Women's or Men's race
Dim iNumPage	' Number of pages
' Get parameters
iPageSize = 20
' Retrieve page to show or default to 0
If Request.QueryString("page") = "" Then
   iRecCurrent = 0	' First Record
Else
   iRecCurrent = CInt(Request.QueryString("page"))
End If
' Assign value to race
If gend = "Male" then
	sGender = "Men's"
else
	sGender = "Women's"
End IF
' ****** End Added July 5, 2000 ********
'	SQL statement
sSQL = "SELECT * FROM 5KResults WHERE Gender='"
sSQL = sSQL & gend & "' ORDER BY Time"
Set RecSet = Server.CreateObject("ADODB.Recordset")
RecSet.Open sSQL,"DSN=chiledadsn",adOpenForwardOnly,adLockReadOnly
'****** Added July 5, 2000 *****************
' Get the count of the records
Do while not RecSet.EOF
	J = J + 1
	RecSet.MoveNext
Loop
iRecEnd = J -1
' Get the number of pages
iNumPage = CInt(iRecEnd/iPageSize)
' If the request page falls outside the acceptable range,
' give them the closest match (0 or max)
If iRecCurrent > iRecEnd Then iRecCurrent = iRecEnd
If iRecCurrent < 0 Then iRecCurrent = 0
If iRecCurrent < iRecEnd Then
   iRecNext = iRecCurrent + iPageSize
Else
	iRecNext = iRecEnd
End If
If iRecCurrent > 0 Then
	iRecPrev = iRecCurrent - iPageSize
Else
	iRecPrev = 0
End If
' Do this so when calling the las page we only loop through
' the number of records we have if less than the iPageSize
if (iRecNext - iRecEnd ) > 0 Then
	iRecMax = iRecEnd - iRecCurrent
Else
	iRecMax = iPageSize
End If
'********End Added July 5, 2000 ********
' Start at the beginning of the database
RecSet.MoveFirst
'Move to the record we want to start at
RecSet.Move(iRecCurrent)
' use this when creating links
' doesn't matter what this page is named
strScriptName = Request.ServerVariables("SCRIPT_NAME")
%>
<%
Sub NavBar()
	Dim iPage
	Dim iVue
	Dim	sNumbers
	Dim sPrev
	Dim sNext
	Dim sFirst
	Dim sLast
	Dim sNavBar
	Dim iLastPage
	iLastPage = iRecEnd - iPageSize
	For i = 0 to (iNumPage - 1)
		iPage = i * iPageSize
		iVue = i + 1
		sNumbers = sNumbers & NavLink(strScriptName,iPage,gend,iVue)
	Next
	if iRecCurrent <> 0 Then
		sFirst = NavLink(strScriptName,0,gend,"First")
		sPrev = NavLink(strScriptName,iRecPrev,gend,"Previous")
	End If
	If (iRecCurrent + iRecMax) < iRecEnd Then
		sNext = NavLink(strScriptName,iRecNext,gend,"Next")
		sLast = NavLink(strScriptName,iLastPage,gend,"Last")
	End If
	sNavBar = sNumbers & "<BR>" & sFirst & sPrev & sNext & sLast
	Response.Write(sNavBar)
End Sub
%>
<%
' Creates the link used by the navigation sub
Function NavLink(scriptName,pageNum,gendr, sWord)
	Dim strLink
	strLink = strLink & "<a HREF='"
	strLink = strLink & scriptName
	strLink = strLink & "?page="
	strLink = strLink & pageNum
	strLink = strLink & "&gender="
	strLink = strLink & gendr
	strLink = strLink & "'>"
	strLink = strLink & sWord
	strLink = strLink & "</a>&nbsp;&nbsp;"
	NavLink = strLink
End Function
%>
<html>
   <head>
	<title>5K Race Results </title>
	<meta name="description" content="An example of paging through a database.">
	<meta name="keywords" content="Active Server Pages, ASP, database, paging">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<base target="_top">
	<meta name="language" content="en-us">
	<meta name="robots" content="INDEX">
	<meta name="revisit-after" content="14 days">
	<meta http-equiv="pragma" content="no-cache">
   </head>
<body>
<!-- Database Table -->
<h3><% =sGender %> 5K Race</h3>
<p><strong>Records</strong>: <% =iRecCurrent %> - <% = iRecCurrent + iRecMax %> of <% =iRecEnd %></p>
<p><% NavBar %></p>
<%
' Use these for debugging
'Response.Write ("iRecCurrent: " & iRecCurrent & "<BR>")
'Response.Write("iRecEnd: " & iRecEnd & "<BR>")
'Response.Write("iRecMax: " & iRecMax & "<BR>")
'Response.Write("iRecNext: " & iRecNext & "<BR>")
'Response.Write("iRecPrev: " & iRecPrev & "<BR>")
'Response.Write(CInt(iRecEnd/iPageSize) & "<BR>")
%>
<table border="0" cellPadding="1" cellSpacing="0" width="425px">
   <tr bgColor="blue">
	<td style="WIDTH: 130px" width="150" bgcolor="#388C40"><strong>Name</strong></td>
	<td style="WIDTH: 35px" width="35" bgcolor="#388C40"><strong>Age</strong></td>
	<td style="WIDTH: 90px" width="150" bgcolor="#388C40"><strong>City</strong></td>
	<td style="WIDTH: 35px" width="45" bgcolor="#388C40"><strong>State</strong></td>
	<td style="WIDTH: 50px" width="75" bgcolor="#388C40"><strong>Time</strong></td>
	<td style="WIDTH: 50px" width="75" bgcolor="#388C40"><strong>Pace</strong></td></tr>
<%
For i = 0 to iRecMax
  	if i mod 2 then
	Response.write ("<TR bgColor=""#008080""><TD>")
  	else
  		Response.Write("<TR><TD>")
  	end if
	Response.Write(RecSet("FirstName") & " ")
	Response.Write(RecSet("LastName")& "</TD>")
	Response.Write("<TD>" & RecSet("age") & "</TD>")
	Response.Write("<TD>" & RecSet("City") & "</TD>")
	Response.Write("<TD>" & RecSet("State") & "</TD>")
	Response.Write("<TD>" & RecSet("Time" )& "</TD>")
	Response.Write("<TD>" & RecSet("Pace") & "</TD>")
	Response.Write("</TR>")
' Move to the next record
  	RecSet.MoveNext
Next
' Clean up after yourself
	RecSet.Close
	Set RecSet = Nothing
%>
	</table>
<p><% Call NavBar %></p>
<!-- End Database Table -->
	</body>
</html>
```

