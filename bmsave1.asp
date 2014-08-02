<!--#include file="conn.asp"-->
<%
lid=request.QueryString("id")
fgld=request.form("fgld")
dwbm=request.form("dwbm")



if fgld<>"" and dwbm<>"" then
sql="select * from dw where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then
rs("dwbm")=dwbm
rs("fgld")=fgld
rs.update
response.redirect "bmgl.asp"
end if
else 
   response.write "<script language=JavaScript>{window.alert('分管领导和单位序号不允许为空');window.location.href='bmgl.asp';}</script>"

end if
%>
