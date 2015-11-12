<head>
<!-- #INCLUDE FILE="checklogin.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<title>Status</title>
</head>
<!-- #INCLUDE FILE="nav.asp" -->

<%
filename = "status.asp"
filenameaction = "editstatus.asp"
keyname = "status"
tablename = "tblstatus"
%>

<p><a href="<%= filename %>?action=1">create a new <%= keyname %></a></p>

<%
action = request.querystring("action")

'insert a new category
if action = 1 then
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
	mySQL = "INSERT INTO " & tablename & "(" & keyname & ") VALUES (""new " & keyname & """);"
	conn.execute(mySQL)
	closeconn()
	response.redirect(filename)
end if

'display users
set Conn = server.createobject("adodb.connection")
Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "Select * from " & tablename & " where ID > 1;"
'Response.Write strSQL
'response.end
rs.Open strSQL, conn,1
count = rs.recordcount
if count <> 0 then
	data = rs.GetRows()
	'Data is retrieved so close all connections
	closers()
	closeconn()
	'Setup for array usage
	iRecFirst   = LBound(data, 2)
	iRecLast    = UBound(data, 2)
%>
<%=tablehead%>
<tr>
	<th><%= keyname %></th>
</tr>
<%
For I = iRecFirst To iRecLast
	id 			= data(0,I)
	category 	= data(1,I)
%>
<tr>
	<td><a href="<%= filenameaction %>?id=<%=id%>"><%= category %> (<%=id%>)</a>&nbsp;</td>
</tr>
<%
	id 			= ""
	category 		= ""
Next
response.write "</table>"
end if
response.write "<p>users: " & count & "</p>"
%>