
<%


dim objConn
dim sConnect
set objConn=Server.CreateObject("ADODB.connection")
sConnect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("movie.mdb")
objConn.ConnectionString=sConnect
objConn.open
%>