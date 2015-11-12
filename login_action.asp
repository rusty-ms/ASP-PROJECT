<%Session.Timeout=120%>
<!-- #INCLUDE FILE="common.asp" -->
<%
' Intercept all exceptions to display user-friendly error
'On Error Resume Next 

if Request.Form("email") = "" then
	Response.Redirect("login.asp?l=email_blank")
elseif Request.Form("mynumber") = "" then
	Response.Redirect("login.asp?l=password_blank")
End If

Dim rs, FormLogin, FormPwd, LoginID, logincount, strSQL
'set variables from form
FormLogin = trim(Request.Form("email"))
FormLogin = replace(FormLogin,"'","")
FormPwd = Trim(Request.Form("mynumber"))
crypt = EncryptText(FormLogin,FormPwd)

'create instance of recordset, and run  query
set Conn = server.createobject("adodb.connection")
Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "Select ID,email,logindate,loginip,mynumber FROM tblusers where active = 2 and email = '" & FormLogin & "' and mynumber = '" & crypt & "';"
Response.Write strSQL
'response.end
rs.Open strSQL, conn, 1,3

'run login or return to login page
if not rs.eof then
	dim email, allowedimages
	crypt = rs("mynumber")
	Response.Cookies("crypt") = crypt
	Response.Cookies("crypt").Expires = Date + 30
	Response.Cookies("stampedname") = FormLogin
	Response.Cookies("stampedname").Expires = Date + 30
	rs("logindate") = date
	rs("loginip") = request.servervariables("REMOTE_ADDR")
	rs.Update

	'SET Cookie for email
	if request.form("rememberme") = 1 then
		Response.Cookies("email") = FormLogin
		Response.Cookies("email").Expires = Date + 14
	end if

else
	closers()
	closeconn()
	Response.Redirect("login.asp?l=incorrect_login_or_password")
End if

'close connections,etc..
closers()
closeconn()

'final redirect 
Response.Redirect "login.asp?loggedin=yes"
%>