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
   ksyj="同意。"
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
set rs3=server.createobject("adodb.recordset")
rs3.open sql,connstr,3,2
if not rs3.eof then 
  ZG1=RS3("ZG")
  SFLFJ1=RS3("SFLFJ")
END IF
RS3.CLOSE



sql="select * from QXJDJb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("FGLDYJ")=ksyj
rs("JLD")=KS
RS("FGLD")=SQR1
rs("fgldrq")=date()

ZW1=rs("QJRZW")
if ks="0"   then
  rs("QJzt")="9"
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请不同意,已终审!');window.location.href='SpJLD.asp';}</script>"
   response.end
END IF

IF KS="1" AND ZW1="1" and rs("SFLFJ") = "NO"  THEN
   rs("QJzt")="7"
   RS("ZGZT")="1"
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请同意,已终审!');window.location.href='SpJLD.asp';}</script>"
   response.end
   END IF 
IF KS="1" AND (ZW1="3" or zw1="6"  or rs("SFLFJ") = "YES" )THEN
   rs("QJzt")="4"
   rs.update
   response.write "<script language=JavaScript>{window.alert('申请人分管领导已审,成功传局长审批!');window.location.href='SpJLD.asp';}</script>"
   response.end
   END IF 
end if
'rs.close

%>
