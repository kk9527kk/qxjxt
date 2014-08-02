<!--#include file="conn.asp"-->
<%
lid=request.QueryString("id")
dw=request.form("dw")
id=request.form("id")
zg=request.form("zg")
mm=request.form("mm")
zw=request.form("zw")
xh=request.form("xh")

if zg="" or mm="" then
   response.write "<script language=JavaScript>{window.alert('不允许账号和密码为空！');window.location.href='zhgl.asp';}</script>"
   response.end
 else
sql="select * from zg where id='"&lid&"'"
response.write(sql)
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then
rs("dw")=dw
rs("zg")=zg
'rs("id")=id
rs("mm")=mm
rs("zw")=zw
rs("xh")=xh
rs.update
response.redirect "zhgl.asp"
end if
end if
set rs=nothing

%>
