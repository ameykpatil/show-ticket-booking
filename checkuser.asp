<%@ Language=VBScript %>
<%option Explicit%>

<!-- #include file="adovbs.inc" -->
<!-- #include file="conn.asp" -->
<%
	Dim objRS
	Set objRS=Server.CreateObject("ADODB.Recordset")
	Set objRS.ActiveConnection=objConn
	objRS.open "select uname,password from Table1 where uname='" & request("u") & "' and password='" & request("p") & "'",objConn, ,adLockOptimistic
	if not objRS.EOF then
		response.redirect("movie.html")
	else
		response.redirect("login.html")
	end if
%>
<%
	Session("username")=Request.Form("u")
    Response.Redirect("home.html")
	objRS.Close
	Set objRS=Nothing
	objConn.Close
	Set objConn=Nothing
%>	
	
	
	
	
	
	
	
	
	
	
	