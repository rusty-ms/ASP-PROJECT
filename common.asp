<style type="text/css">

p {
	font-size: smaller;
}

.nodec{
	text-decoration: none;
}

table {
	font-size: smaller;
}

.red {
	color: red;
}
.thstyle {
	background-color: black;
}
.formfont {
	font-family: arial;
}		


body {
	background: #fafafa url(http://jackrugile.com/images/misc/noise-diagonal.png);
	color: #444;
	font: 100%/30px 'Helvetica Neue', helvetica, arial, sans-serif;
	text-shadow: 0 1px 0 #fff;
	margin: 0;
	padding: 0;
	height: 100%;
	width: 100%;
	text-align:center;
}

/* unvisited link */
a:link {
    color: #000;
}

/* visited link */
a:visited {
    color: #000;
}

/* mouse over link */
a:hover {
    color: #000;
}

/* selected link */
a:active {
    color: #000;
}

strong {
	font-weight: bold; 
}

em {
	font-style: italic; 
}

table {
	background: #f5f5f5;
	border-collapse: separate;
	box-shadow: inset 0 1px 0 #fff;
	font-size: 12px;
	line-height: 24px;
	margin: 30px auto;
	text-align: left;
	width: 800px;
}	

th {
	background: url(http://jackrugile.com/images/misc/noise-diagonal.png), linear-gradient(#777, #444);
	border-left: 1px solid #555;
	border-right: 1px solid #777;
	border-top: 1px solid #555;
	border-bottom: 1px solid #333;
	box-shadow: inset 0 1px 0 #999;
	color: #fff;
  font-weight: bold;
	padding: 10px 15px;
	position: relative;
	text-shadow: 0 1px 0 #000;
text-align: center;
color: white;	
}

th:after {
	background: linear-gradient(rgba(255,255,255,0), rgba(255,255,255,.08));
	content: '';
	display: block;
	height: 25%;
	left: 0;
	margin: 1px 0 0 0;
	position: absolute;
	top: 25%;
	width: 100%;
}

th:first-child {
	border-left: 1px solid #777;	
	box-shadow: inset 1px 1px 0 #999;
}

th:last-child {
	box-shadow: inset -1px 1px 0 #999;
}

td {
	border-right: 1px solid #fff;
	border-left: 1px solid #e8e8e8;
	border-top: 1px solid #fff;
	border-bottom: 1px solid #e8e8e8;
	padding: 10px 15px;
	position: relative;
	transition: all 300ms;
}

td:first-child {
	box-shadow: inset 1px 0 0 #fff;
}	

td:last-child {
	border-right: 1px solid #e8e8e8;
	box-shadow: inset -1px 0 0 #fff;
}	

tr {
	background: url(http://jackrugile.com/images/misc/noise-diagonal.png);	
}

tr:nth-child(odd) td {
	background: #f1f1f1 url(http://jackrugile.com/images/misc/noise-diagonal.png);	
}

tr:last-of-type td {
	box-shadow: inset 0 -1px 0 #fff; 
}

tr:last-of-type td:first-child {
	box-shadow: inset 1px -1px 0 #fff;
}	

tr:last-of-type td:last-child {
	box-shadow: inset -1px -1px 0 #fff;
}	

tbody:hover td {
	color: transparent;
	text-shadow: 0 0 3px #aaa;
}

tbody:hover tr:hover td {
	color: #444;
	text-shadow: 0 1px 0 #fff;
}

#navigation {
  background-color: #eee;
}
#navigation ul {
  margin: 0;
  padding: 0; 
}
#navigation li {
  border-right: 1px solid #ddd;
  display: block;
  float: left;
  margin: 0;
}
#navigation li:last-child {
  border-right-width: 0;
}
#navigation a {
  background-color: #eee;
  color: #333;
  display: block;
  padding: .75em 1.5em;
  text-decoration: none;
  transition: all .25s ease-in-out;
  -moz-transition: all .25s ease-in-out;
  -webkit-transition: all .25s ease-in-out;
}
#navigation a:hover {
  background-color: #ddd;
}
.clearfix {
  *zoom: 1;
}
.clearfix:after, 
.clearfix:before {
  content: '';
  display: table;
}
.clearfix:after {
  clear: both;
}	
</style>
<%

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

dbPath = Server.MapPath("projectdb.mdb")

tablehead = "<table class=""rwd-table"">"
'tablehead = "<table border=""1"" cellspacing=""0"" cellpadding=""4"" bordercolor=""silver"">"

sub openconn()
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
end sub

sub closers
	rs.Close
	set rs = nothing
end sub

sub closeconn
	conn.close
	set conn = nothing
end sub

Function RandGen
	Set rfs = CreateObject("Scripting.FileSystemObject")
	RandGen = rfs.GetBaseName(rfs.GetTempName)
	set rfs = Nothing
	RandGen = Right(RandGen, Len(RandGen) - 3)
End Function


Function yesno(cdata,cname)
	if cdata = "on" then
		yesno = "<input type='checkbox' name='" & cname & "' checked>"
	else
		yesno = "<input type='checkbox' name='" & cname & "'>"
	end if
end function

Function MakeCombo(comboid,tabletouse,fieldtouse)
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
    set rs=Server.CreateObject("ADODB.recordset")
    rs.Open "Select * from " & tabletouse & ";", conn 'change your table name .
    do while not rs.EOF ' before this you can check for eof or bof
		data = rs(fieldtouse)
		dataid = rs("id")
			if comboid = data then
				ckd = " selected"
			end if
        Response.Write vbtab & "<Option" & ckd & " value='"& dataid & "'>" &(rs(fieldtouse)) &"</option>" & vbcrlf
		ckd = ""
        rs.MoveNext()
    loop
	closers()
	closeconn()
End Function

FUNCTION PreSubmit2(p_sTargetString)
     PreSubmit2 = REPLACE(p_sTargetString,"textarea","")
     PreSubmit2 = REPLACE(PreSubmit2,"'","&#39;")
     PreSubmit2 = REPLACE(PreSubmit2,"""","&quot;")
     PreSubmit2 = REPLACE(PreSubmit2,"<","&lt;")
     PreSubmit2 = REPLACE(PreSubmit2,">","&gt;")
	 PreSubmit2 = Trim(PreSubmit2)
END FUNCTION

Function counttherows(whattable,whatkey)
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
    set rs=Server.CreateObject("ADODB.recordset")
	mainSQL = "Select " & whatkey & " FROM " & whattable & ";"
	response.write mysql
	rs.Open mainSQL, Conn, 1
	counttherows = rs.recordcount
	rs.Close
	set rs = nothing
	conn.close
	set conn = nothing
End function

Function counttherows2(mainsql)
	set Conn = server.createobject("adodb.connection")
	Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
    set rs=Server.CreateObject("ADODB.recordset")
	rs.Open mainsql, Conn, 1
	counttherows2 = rs.recordcount
	rs.Close
	set rs = nothing
	conn.close
	set conn = nothing
End function

Private Function EncryptText(ByVal strEncryptionKey, ByVal strTextToEncrypt)
    ' Declare variables
    Dim outer, inner, Key, strTemp
    ' For each character in strEncryptionKey
    For outer = 1 To Len(strEncryptionKey)
        ' Get a character to use as our encryption 
        ' key in this iteration of the OUTER loop
        key = Asc(Mid(strEncryptionKey, outer, 1))
        ' For each character in strTextToEncrypt
        For inner = 1 To Len(strTextToEncrypt)
            ' Update our encrypted text
            strTemp = strTemp & Chr(Asc(Mid(strTextToEncrypt, inner, 1)) Xor key)
            ' Change our encryption key to mix things up in the INNER loop.
            key = (key + Len(strEncryptionKey)) Mod 256
        Next
        ' Update the strTextToEncrypt variable before 
        ' the next iteration of the OUTER loop
        strTextToEncrypt = strTemp
        ' Reset strTemp for the next iteration of the OUTER loop.
        strTemp = ""
    Next
    ' Assign the value of the encrypted text to the function name 
    ' so we can return the value to the caller
    EncryptText = strTextToEncrypt
End Function
%>