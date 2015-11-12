<%
crypt = ""
Response.Cookies("crypt") = " "
Response.Cookies("crypt").Expires = Date() - 1
Response.Cookies("stampedname") = " "
Response.Cookies("stampedname").Expires = Date() - 1
response.redirect("login.asp")
%>
