<%@ Language=VBScript %>
<!-- #include file="adovbs.inc" -->
<!-- #include file="conn.asp" -->
<%

	User = Request.Form("u")	
	Pass = Request.Form("p")
	response.write "Username=" & User
	response.write "Password=" & Pass

	
%>


	

