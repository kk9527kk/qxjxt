<% 
dim db,conn,connstr
db="txl.mdb"
set Conn = server.CreateObject("ADODB.Connection")
connstr="provider=microsoft.jet.oledb.4.0;data source="& server.MapPath(""&db&"")
conn.Open connstr
%>