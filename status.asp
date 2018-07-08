<%@ Language=VBScript %>
<!-- #include file="adovbs.inc" -->
<%
	dim objConn
	dim sConnect
	set objConn=Server.CreateObject("ADODB.connection")
	sConnect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("status.mdb")
	objConn.ConnectionString=sConnect
	objConn.open
	dim oRs
	Set oRs=Server.CreateObject("ADODB.Recordset")
	Set oRs.ActiveConnection=objConn
	oRs.open "select seatname from Table1 where status=' ' ",objconn, ,adLockOptimistic,adCmdTable
	do until oRs.EOF
 for each x in oRs.Fields
 Response.Write(x.value&"<br/>")
 next
oRs.MoveNext
loop

oRs.close
objConn.close
		

%>