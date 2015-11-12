<head>
<!-- #INCLUDE FILE="checklogin.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<%
strFormList = Request.Form("projorder")
strIDList = request.form("id")
if strFormList = "" then
	Response.Write "<p>Error</p>"
	response.redirect("projects.asp?order=error")
Else

	'on Error resume next
	FormList = split(strFormList, ",")
	IDList = split(strIDList, ",")

	For iLoop = LBound(FormList) to UBound(FormList)

		pageorder = Trim(FormList(iLoop))
		'make sure value is a number
			if not IsNumeric(pageorder) then
			pageorder = "1"
			end if

		id = Trim(IDList(iLoop))
			'make sure value is a number
			if not IsNumeric(id) then
			response.write "<p>Error, not a valid number</p>"
			response.end
			end if

		if pageorder > 0 then
		set Conn = server.createobject("adodb.connection")
		Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
		mySQL = "UPDATE tblproj SET projorder = " & pageorder & " WHERE ID = " & ID & ";"
		conn.execute(mySQL)
		end if

	Next

end if

response.redirect("projects.asp?order=complete")
%>

<p><a href="projects.asp">Complete!</a></p>