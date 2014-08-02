<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
lid=request.QueryString("id")
'response.write(lid)
if lid="" or  request.Cookies("adminuser")="" then
   response.write "来源未知或数据丢失!"
   response.end
end if

zxyj=request.form("zxyj")
zx=request.form("zx")

if zxyj="" and zx="1" then
   zxyj="申请情况属实,提交局领导审批"
elseif ksyj="" and zx="0" then
   zxyj="情况不属实,,请重申请"
else
   zxyj=zxyj
end if


sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zx1=rs1("zg")
end if
rs1.close


sql="select * from sqspb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("zxyj")=zxyj
rs("zx")=zx1
rs("zxrq")=date()
 if zx="1" then
   rs("zt")="3"
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请已经成功提交到局领导!');window.location.href='Sp1.asp';}</script>"
   response.end
 else
   rs("zt")="6"
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请不同意,已终审!');window.location.href='Sp1.asp';}</script>"
   response.end
 end if
end if
 


%>
