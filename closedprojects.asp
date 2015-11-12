<head>
<!-- #INCLUDE FILE="checklogin.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<title>Projects</title>
</head>
<!-- #INCLUDE FILE="nav.asp" -->

<p><a href="projects.asp?action=1">create a new project</a></p>

<%
stampedname = Request.Cookies("stampedname")
sortby = trim(request.querystring("sortby"))
again = trim(request.querystring("again"))

if again <> "" and again = sortby then
	sortby = sortby & " desc"
end if

if sortby = "" then
	sortby = "projorder"
	again = ""
end if

action = request.querystring("action")

'insert a new project
if action = 1 then
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
	mySQL = "INSERT INTO tblProj(ProjName,ProjDate,ProjPrivate,projlastuser,projrand) VALUES ('**New Project**',#" & date() & "#,'on','" & stampedname & "','" & randgen() & "');"
	conn.execute(mySQL)
	closeconn()
	response.redirect("projects.asp")
end if

'display projects
set Conn = server.createobject("adodb.connection")
Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "Select tblproj.id,tblproj.projname,tblproj.projdate,tblproj.projdesc,tblcat.category,tblstatus.status,tblproj.projorder, tblproj.projlastuser,tblproj.projlastdate,tblproj.projprivate,tblproj.projlastuser from tblproj,tblcat,tblstatus where tblproj.projcategory = tblcat.id and tblproj.projstatus = tblstatus.id and tblproj.projstatus = 3 order by tblproj." & sortby & ";"
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
<form action="setorder.asp" method="post">
<%=tablehead%>

<tr>
	<th><a href="projects.asp?sortby=projorder&again=<%= sortby %>"><font color="white">Order</font></a></th>
	<th><a href="projects.asp?sortby=projname&again=<%= sortby %>"><font color="white">Name</font></a></th>
	<th><a href="projects.asp?sortby=projdate&again=<%= sortby %>"><font color="white">Date</font></a></th>
	<th><a href="projects.asp?sortby=projcategory&again=<%= sortby %>"><font color="white">Category</font></a></th>
	<th><a href="projects.asp?sortby=projstatus&again=<%= sortby %>"><font color="white">Status</font></a></th>
	<th><a href="projects.asp?sortby=projlasthate&again=<%= sortby %>"><font color="white">Last</font></a></th>
</tr>
<%
For I = iRecFirst To iRecLast
	id 			= data(0,I)
	projname 	= data(1,I)
	projdate 	= data(2,I)
	projlastdate= data(8,I)
	category	= data(4,I)
	status 		= data(5,I)
	projorder 	= data(6,I)
	projprivate	= data(9,I)
	projlastuser= data(10,I)
	if projprivate = "on" and projlastuser <> stampedname then
	else
%>
<tr>
	<td><input type="text" name="projorder" value="<%= projorder %>" size="1" maxlength="4">
		<input type="hidden" name="id" size="2" value="<%= id %>">&nbsp;</td>
	<td><a href="editproject.asp?id=<%=id%>"><%= projname %> (<%=id%>)</a>&nbsp;</td>
	<td><%= projdate %>&nbsp;</td>
	<td><%= category %>&nbsp;</td>
	<td><%= status %>&nbsp;</td>
	<td><%= projlastdate %>&nbsp;</td>
</tr>
<%
	end if
	id 			= ""
	projname 	= ""
	projdate 	= ""
	projlastdate= ""
	category	= ""
	status		= ""
	projorder	= ""
	projprivate	= ""
Next
%>
	<tr bgcolor="#CCCCFF">
		<td colspan="6" align="left"><input type="submit" name="submit" value="update order"></td>
	</tr>
</table>
</div>
</form>
<%
end if
response.write "<p><small>total projects: " & count & "<br>" & vbcrlf & _
				"current user: " & Request.Cookies("stampedname") & "<br>" & vbcrlf & _
				"date/time: " & Now() & "</small></p>" 
%>