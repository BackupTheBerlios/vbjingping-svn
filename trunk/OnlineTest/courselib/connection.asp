
<%
	'����������������ASP����
	dbfile="./data/db1.mdb"
	dim conn,rs
	set conn=Server.CreateObject("adodb.connection")
	set rs=Server.CreateObject("adodb.recordset")
	connstr="driver={microsoft access driver (*.mdb)};dbq=" & server.mappath(""& dbfile &"")
	conn.open connstr
'	response.Write conn.State
	
%>
