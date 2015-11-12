<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!-- #INCLUDE FILE="common.asp" -->
<html>
<head>
	<title>Password Helper</title>
</head>
<body>

<%
'if you forget your password, use this to create a new password to enter directly into the database.
'enter your username (case sensitive) and new password, then run page to display password.
'then open the database and replace the value in the field mynumber with the displayed password.
'username is case sensitive.

current_user_name = ""
new_password = ""

data = EncryptText(current_user_name,new_password)

if data <> "" then
response.write "<p>new password=<input type='text' value=" & data & "'></p>"
else
%>
<hr>
<p>To use this file, please open with a text editor to read instructions!</p>

<% end if %>
</body>
</html>
