<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
'response.write(uid)

lid=request.QueryString("id")
'response.write(lid)
if lid="" or  request.Cookies("adminuser")="" then
   response.write "来源未知或数据丢失!"
   response.end
end if

ksyj=request.form("ksyj")
ks=request.form("KS")

if ksyj="" and ks="1"  then
   ksyj="同意请假。"
elseif ksyj="" and ks="0" then
   ksyj="不同意请假申请。"
else
   ksyj=ksyj
end if

sql="select * from ZG where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    sqr1=rs1("zg")
end if
rs1.close





sql="select * from QXJDJb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("jzYJ")=ksyj
rs("ld")=KS
RS("jz")=SQR1
rs("jzrq")=date()
if ks="0"   then
  rs("QJzt")="9"
     rs.update
   response.write "<script language=JavaScript>{window.alert('申请不同意,已终审!');window.location.href='Spjz.asp';}</script>"
   response.end
else
   rs("qjzt")="5"
   rs("zgzt")="1"
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请同意,已终审!');window.location.href='Spjz.asp';}</script>"
   response.end
END IF 
end if

%>
