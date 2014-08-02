<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='index.htm';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
if uid<>"" then
lid=request("id")
sql="select * from sqspb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then
rs.delete
rs.update
response.redirect "sp1.asp"
end if
else
response.redirect "sp1.asp"
end if
%>
