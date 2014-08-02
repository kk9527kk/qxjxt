<!--#include file="conn.asp"-->
<%
lid=request("id")
sql="select * from zg where id='"&lid&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then
rs.delete
rs.update
response.redirect "zhgl.asp"
end if
%>
