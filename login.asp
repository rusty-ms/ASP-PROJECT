<%
<!-- #INCLUDE FILE="common.asp" -->

'if already logged in then skip page.
If Request.Cookies("crypt") = "" Then

Dim ntuser 
ntuser = Request.ServerVariables("LOGON_USER") & ""
Set objDomain = GetObject ("GC://RootDSE")
objADsPath = objDomain.Get("defaultNamingContext")
Set objDomain = Nothing
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.provider ="ADsDSOObject"
objConn.Properties("User ID") = "DOMAINUSER" 'domain account with read access to LDAP
objConn.Properties("Password") = "DOMAINPASSWORD" 'domain account password
objConn.Properties("Encrypt Password") = True
objConn.open "Active Directory Provider"
Set objCom = CreateObject("ADODB.Command")
Set objCom.ActiveConnection = objConn
ntuser = Replace(ntuser,"DOMAINNAME\", "")
objCom.CommandText ="select name,department FROM 'GC://"+objADsPath+"' where sAMAccountname='"+ntuser+"' ORDER by sAMAccountname"
'=======Execute queury on LDAP for all accounts=========
Set objRS = objCom.Execute
Dim NAME 
NAME = objRS("name")
Response.Cookies("stampedname") = NAME
Response.Cookies("crypt") = NAME
Response.Cookies("crypt").Expires = Date + 30
Response.Cookies("stampedname").Expires = Date + 30
response.redirect("projects.asp")

else


response.write "Already Logged in."
response.redirect("projects.asp")
response.write "<p><a href='logout.asp'>logout</a>"
'response.redirect ("/myaccount/")


end if
%>