<head>
<!-- #INCLUDE FILE="checklogin.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<title>Projects</title>
</head>
<!-- #INCLUDE FILE="nav.asp" -->

<%
filename = ""
filenameaction = ""
keyname = "password"
tablename = ""

'update project if form is sent.
formid = request.form("formid")
mynumber = Request.Form("mynumber")

if formid <> "" then

	if mynumber = "" then
		response.write "<p class='red'>error, password is blank!</p>" & vbcrlf & _
		"<p><a href='editpassword.asp?id=" & formid & "'>try again</a></p>"
		response.end
	end if

	'id = request.form("formid")
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
	Set rs = Server.CreateObject("ADODB.RecordSet")
	mainSQL = "Select * from tblusers where id = " & formid & ";"
	'response.write mainsql
	rs.Open mainSQL, Conn, 1, 3

	FormLogin = trim(Request.Form("email"))
	FormLogin = replace(FormLogin,"'","")
	FormPwd = Trim(mynumber)
	crypt = EncryptText(FormLogin,FormPwd)

	'update database
	rs("mynumber") = crypt

	'set new cookie
	Response.Cookies("crypt") = crypt
	Response.Cookies("crypt").Expires = Date + 30

	rs.Update
	closers()
	closeconn()
	response.redirect("editusers.asp?id=" & formid & "&updated=3")
end if

id = request.querystring("id")
'display user
set Conn = server.createobject("adodb.connection")
Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "Select * from tblusers where id = " & id & ";"
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
end select
%>

<p><strong>edit <%= keyname %></strong></p>

<form action="editpassword.asp" method="post">
<%=tablehead%>
<tr>
	<td><strong>name</strong></td>
	<td valign="top"><%= email %>&nbsp;</td>
</tr>

<tr>
	<td><strong>new password</strong></td>
	<td valign="top"><input type="password" name="mynumber">&nbsp;</td>
</tr>

</table>
<input type="hidden" name="email" value="<%=email%>">
<input type="hidden" name="formid" value="<%=id%>">
<p><input type="submit" value="update" onclick="return confirm('You are about to change your password. Continue?')"></p>
</form>