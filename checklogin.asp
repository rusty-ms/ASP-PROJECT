<%
Response.Expires = 60
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

If Request.Cookies("stampedname") = "" Then
	Response.Redirect("login.asp?message=not_logged_in")
else
	crypt = Request.Cookies("crypt")
end if
%>