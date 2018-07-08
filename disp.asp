<%@ Language=VBScript %>
<!-- #include file="adovbs.inc" -->
<%
	
	dim objConn
	dim sConnect
	set objConn=Server.CreateObject("ADODB.connection")
	sConnect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("seats.mdb")
	objConn.ConnectionString=sConnect
	objConn.open
	dim oRs
	Set oRs=Server.CreateObject("ADODB.Recordset")
	Set oRs.ActiveConnection=objConn
	oRs.open "Table1",objconn, ,adLockOptimistic,adCmdTable
	a=request("movielist")
	b=request("th")
	c=request("show")
	d=request("date")
	if d="tod" then
	e=date()
	oRs.AddNew
	
		oRs("moviename")=a
		oRs("show")=c
		oRs("theatre")=b
		oRs("date")=e
		oRs.Update
	else
	f=dateadd("d",1,date())
	oRs.AddNew
	
		oRs("moviename")=a
		oRs("show")=c
		oRs("theatre")=b
		oRs("date")=f
		oRs.Update
	end if
	response.redirect"seats.html"
	oRs.close
	objConn.close
%>