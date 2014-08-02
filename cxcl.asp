<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
'lid=request.QueryString("id")
uid=request.Cookies("adminuser")
'response.write(uid)
sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw1=rs1("dw")
    zg1=rs1("zg")
    zw1=rs1("zw")
    
end if
rs1.close

%>


<%
rqq=request.form("T1")
rqz=request.form("T2")

if not (isdate(rqq) and  isdate(rqz)) then
   response.write "<script language=JavaScript>{window.alert('请输入正确的日期格式!');window.location.href='tycx.asp';}</script>"
   response.end
end if 

    
if DateDiff("d",rqq,rqz)<0 then 
   response.write "<script language=JavaScript>{window.alert('请输入起时间大于输入时间止时间！');window.location.href='tycx.asp';}</script>"
   response.end
end if

    response.cookies("rqq") = rqq
    response.cookies("rqz") = rqz

    response.cookies("flag") = "loginok"
    response.cookies("adminuser") = uid
'    rqq=request.Cookies("rqq")
 '   rqz=request.Cookies("rqz")

'rqq1=FormatDateTime(rqq,1)
'rqz1=FormatDateTime(rqz,1)
Response.Redirect "tycxcl.asp"

%>