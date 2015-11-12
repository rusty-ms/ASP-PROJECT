<head>
<!-- #INCLUDE FILE="checklogin.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
</head>
<%
dim user, ID, comment
comment=Request.Form("comment")
user=Request.Form("user")
ID=Request.Form("ID")

Dim strDbConnection3
Dim objConn3
Dim objRS3
Dim sql
strDbConnection3 = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
Set objConn3 = Server.CreateObject("ADODB.Connection")
objConn3.Open(strDbConnection3)
'sql="INSERT INTO tblcomments(TicketID,Comment,User) VALUES ("& ID & "','" & comment & "','" & user & "');"
strSQL = "INSERT INTO tblcomments ([TicketID], [Comment], [User]) VALUES ('" & ID & "', '" & comment & "', '" & user & "')"
response.write strSQL
Set objRS3 = objConn3.Execute(strSQL)

Response.redirect("editproject.asp?id="&ID)
closeconn()





%>