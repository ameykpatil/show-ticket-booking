<html>
<HEAD>
<link rel ="stylesheet" type="text/css" href="images.css"/>
<TITLE>bookmyshow/home</TITLE>
</HEAD>
<body background="pictures/black.jpg">

<img src = "pictures/finalmasks.jpg"  width = "15%" height="20%">
<img src = "pictures/b.jpg"  width = "9%" height="20%">
<img src = "pictures/bms10.gif" width="50%" height = "20%">
<img src = "pictures/b.jpg"  width = "9%" height="20%">
<img src = "pictures/finalmovie.jpg"  width = "15%" height="20%">

<br><br>

<table bgcolor="black" border="0" bordercolor= "white" width="100%">
	<tr>

	<td><a href="home.html"><img src="buttons/home.png" onmouseover= "src = 'buttons/homeover.png'" 	onmouseout= "src = 'buttons/home.png'" height="30" width="150" 	border="0"></a></td>

	<td><a href="movie.html"><img border="0" src="buttons/movie.png" 	onmouseover= "src = 'buttons/moviesover.png'" onmouseout= "src = 	'buttons/movie.png'" height="30" width="150"></a></td>

	<td><a href="reviews.html"><img border="0" src="buttons/reviews.png" 	onmouseover= "src = 'buttons/reviewsover.png'" onmouseout= "src = 	'buttons/reviews.png'" height="30" width="150"></a></td>

	<td><a href="aboutus.html"><img border="0" src="buttons/aboutus.png" 	onmouseover= "src = 'buttons/aboutusover.png'" onmouseout= "src = 	'buttons/aboutus.png'" height="30" width="150"></a></td>


	
	</tr>
</table>
<%@ Language=VBScript %>
<!-- #include file="adovbs.inc" -->
<br>
<br>
<br>
<center>
<fieldset>
<font color="white" size="5">
<%
	dim a
	a=request("credit")
	dim objConn
	dim sConnect
	set objConn=Server.CreateObject("ADODB.connection")
	sConnect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("pay.mdb")
	objConn.ConnectionString=sConnect
	objConn.open
	dim oRs
	Set oRs=Server.CreateObject("ADODB.Recordset")
	Set oRs.ActiveConnection=objConn
	oRs.open "Table1",objconn, ,adLockOptimistic,adCmdTable
	if a=420 then
	response.write  "Hello Saras.....Transaction Successfully Completed!!!"
	
	elseif a=150 then
		response.write "Hello Amey.....Transaction Successfully Completed!!!"
	
	elseif a=200 then
		response.write "Hello Nitish.....Transaction Successfully Completed!!!"

	elseif a=100 then
		response.write "Hello Swapnagandha.....Transaction Successfully Completed!!!"
	
	else
	response.write"Please Enter Proper Credit Card Number!!!!"

	end if
	oRs.MoveNext
	oRs.close
	objConn.close
%>
</font>
</center>
</fieldset>
