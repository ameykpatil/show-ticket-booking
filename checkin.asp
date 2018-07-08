<%@ Language=VBScript %>
<%option Explicit%>
<!-- #include file="adovbs.inc" -->
<!-- #include file="conn.asp" -->
<%

	'collect information'

	'connect to database
	dim oRs
	Set oRs=Server.CreateObject("ADODB.Recordset")
	Set oRs.ActiveConnection=objConn
	oRs.open "Table1",objconn, ,adLockOptimistic,adCmdTable
	oRs.AddNew
	
		oRs("fname")=request("first")
		oRs("lname")=request("last")
		oRs("uname")=request("usr")
		oRs("password")=request("pass")
		oRs("gender")=request("gender")
		oRs("street")=request("add1")
		oRs("area")=request("add2")
		oRs("city")=request("city")	
		oRs("pin")=request("pin")
		oRs("state")=request("state")
		oRs("phone")=request("tel")
		oRs("email")=request("email")
	oRs.Update
	'close connection
				oRS.Close
				Set oRS = Nothing 
			objConn.Close
			Set objConn = Nothing
		
%>

 <html>
<body>
</body>
</html>
