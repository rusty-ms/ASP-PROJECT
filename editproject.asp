<head>
<!-- #INCLUDE FILE="checklogin.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<title>Projects</title>
</head>
<!-- #INCLUDE FILE="nav.asp" -->

<%
filename = "projects.asp"
filenameaction = "editproject.asp"
keyname = "projects"
tablename = "tblproj"

'delete project
if request.querystring("delete") = 1 then
	id = trim(request.querystring("id"))
	projrand = trim(request.querystring("projrand"))
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
	mySQL = "Delete FROM " & tablename & " where id = " & id & " and projrand = '" & projrand & "';"
	conn.execute(mySQL)
	closeconn()
	response.redirect(filename)
end if

'update project if form is sent.
if request.form("formid") <> "" then
	id = request.form("formid")
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
	Set rs = Server.CreateObject("ADODB.RecordSet")
	stampedname = Request.Cookies("stampedname")
	
	strSQL = "UPDATE tblproj SET projname = '"& request.form("projname") &"',projdate = '"& request.form("projdate") &"',projdesc = '"& request.form("projdesc") &"',projcategory = '"& request.form("projcategory") &"',projstatus = '"& request.form("projstatus") &"',projlastdate = '"& date() &"',projlastuser = '"& stampedname &"' WHERE id = " & request.form("formid")

	response.write strSQL

	Set objRS = Conn.Execute(strSQL)

	response.redirect("editproject.asp?id=" & id & "&updated=" & updated & "&time=" & time())
end if

id = request.querystring("id")
showprojlastuser = Request.Cookies("stampedname")

'display project
set Conn = server.createobject("adodb.connection")
Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "Select tblproj.id,tblproj.projname,tblproj.projdate,tblproj.projdesc,tblcat.category,tblstatus.status,tblproj.projorder,tblproj.projlastuser,tblproj.projlastdate,tblproj.projprivate,tblproj.projrand from tblproj,tblcat,tblstatus where tblproj.projcategory = tblcat.id AND tblproj.projstatus = tblstatus.id AND tblproj.id = " & id & " order by tblproj.projorder;"
'Response.Write strSQL
'response.end
rs.Open strSQL, conn,1
count = rs.recordcount
	if count <> 1 then
		closers()
		closeconn()
		response.redirect(filename)
	end if
For fnum = 0 To rs.Fields.Count-1
	execute(rs.Fields(fnum).Name & " = rs(" & CHR(34) & rs.Fields(fnum).Name & CHR(34) & ")")
Next
closers()
closeconn()

'private setting checking
if projprivate = "on" and Request.Cookies("stampedname") <> projlastuser then
	response.write "<p class='red'>This is a private project, you do not have permission to view it.</p>"
	response.end
end if

upd = request.querystring("updated")
select case upd
case "1"
	response.write "<p class='red'>updated</p>"
case "2"
	response.write "<p class='red'>invalid date</p>"
end select
%>
<form action="editproject.asp" method="post">
<%=tablehead%>
<TR COLSPAN="2">
<th COLSPAN="2">Editing Project >> "<%= projname %>"</th>
</tr>
<tr>
	<td><strong>Name</strong></td>
	<td valign="top"><input type="text" name="projname" value="<%= projname %>" size="45" maxlength="100">&nbsp;</td>
</tr>
<tr>
	<td><strong>Date</strong></td>
	<td valign="top"><input type="text" name="projdate" value="<%= projdate %>" size="12">&nbsp;</td>
</tr>
<tr>
	<td valign="top"><strong>Description</strong></td>
	<td valign="top"><textarea cols="45" rows="4" name="projdesc" class="formfont"><%= projdesc %></textarea>&nbsp;</td>
</tr>
<tr>
	<td title="category"><strong>Category</strong></td>
	<td valign="top">
	<select name="projcategory">
    <%call makecombo(category,"tblcat","category")%>
	</select>
	</td>
</tr>
<tr>
	<td title="status"><strong>Status</strong></td>
	<td valign="top">
	<select name="projstatus">
    <%call makecombo(status,"tblstatus","status")%>
	</select>
	</td>
</tr>
<tr>
	<td title="optional sort order critera"><strong>Sort</strong></td>
	<td valign="top" title="optional sort order critera"><input type="text" name="projorder" value="<%= projorder %>" size="3">&nbsp;</td>
</tr>
<tr>
	<td title="only last user can view and edit"><strong>Private</strong></td>
	<td valign="top" title="only last user can view and edit"><%=yesno(projprivate,"projprivate")%>&nbsp;</td>
</tr>
<tr>
	<td title="date record last updated"><strong>Last changed</strong></td>
	<td valign="top" title="date record last updated"><%= projlastdate %>&nbsp;</td>
</tr>
<tr>
	<td title="last user who updated the project"><strong>Last user</strong></td>
	<td valign="top" title="last user who updated the project"><%= projlastuser %>&nbsp;</td>
</tr>
<tr>
<td colspan="2">
<input type="hidden" name="projlastdate" value="<%=now()%>">
<input type="hidden" name="projlastuser" value="<%=showprojlastuser%>">
<input type="hidden" name="projrand" value="<%=randgen()%>">
<input type="hidden" name="formid" value="<%=id%>">
<input type="submit" value="Update" class="buttons">
</td>
</tr>
</table>
</form>
</div>

<form action="addcomment.asp" method="post">
<%=tablehead%>
<tr><th>Comments</th></tr>
<%
Dim strDbConnection
Dim objConn2
Dim objRS2
Dim strSQL2
strDbConnection = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set objConn2 = Server.CreateObject("ADODB.Connection")
objConn2.Open(strDbConnection)
strSQL2 = "SELECT ID,TicketID,Comment,User FROM tblcomments where tblcomments.TicketID = " & id
'response.write strSQL2
Set objRS2 = objConn2.Execute(strSQL2)
If objRS2.EOF Then
  Response.Write("No items found")
Else
  Do While Not objRS2.EOF
    Response.Write("<tr><td><b>" & objRS2.Fields("User").value & "</b> says: " & objRS2.Fields("Comment").value & "</td></tr>")
    objRS2.MoveNext()
  Loop
End If
objRS2.Close()
Set objRS2 = Nothing
objConn2.Close()
Set objConn2 = Nothing
%>
<tr>
	<td valign="top"><textarea cols="45" rows="4" name="comment" class="formfont">Enter comments here...</textarea>&nbsp;</td>
	<input type="hidden" name="user" value="<%=showprojlastuser%>">
	<input type="hidden" name="ID" value="<%=id%>">
</tr>
<tr>
<td>
<input type="submit" value="Add comment" class="buttons">
</td>
</tr>
</table>
</div>
</form>
<br />
<p><a href="editproject.asp?delete=1&id=<%=id%>&projrand=<%=projrand%>" onclick="return confirm('Are you sure you want to delete?')" class='red'>delete the project</a></p>