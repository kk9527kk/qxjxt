<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
lid=request.QueryString("id")
if lid="" or  request.Cookies("adminuser")="" then
   response.write "来源未知或数据丢失!"
   response.end
end if

jldyj=request.form("jldyj")
jld=request.form("jld")

if jldyj="" and jld="1" then
   jldyj="同意信息中心意见。"
elseif ksyj="" and jld="0" then
   jldyj="不同意购买。"
else
   jldyj=jldyj
end if

sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    jld1=rs1("zg")
end if
rs1.close


sql="select * from sqspb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("jld")=jld1
rs("jldyj")=jldyj
rs("jldrq")=date()
 if jld="1" then
   rs("zt")="4"
   rs("qssl")=1
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请事项已经成功终审!');window.location.href='Sp2.asp';}</script>"
   response.end
 else
   rs("zt")="7"
   RS("ZGZT")="1"
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请不同意,已终审!');window.location.href='Sp2.asp';}</script>"
   response.end
 end if
end if
 


%>
