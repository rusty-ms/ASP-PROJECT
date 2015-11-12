<head>
<!-- #INCLUDE FILE="checklogin.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<title>Categories</title>
</head>
<!-- #INCLUDE FILE="nav.asp" -->

<%
filename = "cat.asp"
filenameaction = "editcat.asp"
keyname = "category"
tablename = "tblcat"

'delete category
if request.querystring("delete") = 1 then
	id = request.querystring("id")
		if counttherows2("Select id from tblproj where projcategory = " & id & ";") <> 0 then
			response.write "<p>Unable to delete, projects are using this value.</p>"
			response.end
		end if

	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
	mySQL = "Delete FROM " & tablename & " where id = " & id & ";"
	conn.execute(mySQL)
	closeconn()
	response.redirect(filename)
end if

'update category if form is sent.
if request.form("formid") <> "" then
	id = request.form("formid")
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
	Set rs = Server.CreateObject("ADODB.RecordSet")
	mainSQL = "Select * from " & tablename & " where id = " & id & ";"
	'response.write mainsql
	rs.Open mainSQL, Conn, 1, 3

		For Each strItem In Request.Form
			execute("rs(" & CHR(34) & strItem & CHR(34) & ") = request.form(" & CHR(34) & strItem & CHR(34) & ")")
    	Next

	rs.Update
	closers()
	closeconn()
	response.redirect("" & filenameaction & "?id=" & id & "&updated=" & updated & "&time=" & time())
end if

id = request.querystring("id")
'display projects
set Conn = server.createobject("adodb.connection")
Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "Select * from " & tablename & " where id = " & id & ";"
'Response.Write strSQL
'response.end
rs.Open strSQL, conn,1
count = rs.recordcount
For fnum = 0 To rs.Fields.Count-1
	execute(rs.Fields(fnum).Name & " = rs(" & CHR(34) & rs.Fields(fnum).Name & CHR(34) & ")")
Next
closers()
closeconn()

upd = request.querystring("updated")
select case upd
case "1"
	response.write "<p class='red'>updated</p>"
case "2"
	response.write "<p class='red'>invalid date</p>"
case "3"
	response.write "<p class='red'>password changed</p>"
end select
%>
<p><strong>edit <%= keyname %></strong></p>

<form action="<%= filenameaction %>" method="post">

<%=tablehead%>
<tr>
	<th>name</th>
	<td valign="top"><input type="text" name="<%= keyname %>" value="<%= category %>">&nbsp;</td>
</tr>
</table>
<input type="hidden" name="formid" value="<%=id%>">
<p><input type="submit" value="update"></p>
</form>

<% if counttherows(tablename,"id") > 1 then %>

<p><a href="<%= filenameaction %>?delete=1&id=<%=id%>" onclick="return confirm('Are you sure you want to delete?')" class='red'>delete <%= keyname %></a></p>

<% end if %>