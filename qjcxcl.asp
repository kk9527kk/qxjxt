<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
'lid=request.QueryString("id")
uid=request.Cookies("adminuser")
'response.write(uid)
sql="select * from zg  where id='"&uid&"'"
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
dw=request.form("dw")
zg=request.form("zg")


if not (isdate(rqq) and  isdate(rqz)) then
   response.write "<script language=JavaScript>{window.alert('��������ȷ�����ڸ�ʽ!');window.location.href='qjcx.asp';}</script>"
   response.end
end if 

    
if DateDiff("d",rqq,rqz)<0 then 
   response.write "<script language=JavaScript>{window.alert('��������ʱ���������ʱ��ֹʱ�䣡');window.location.href='qjcx.asp';}</script>"
   response.end
end if

    response.cookies("rqq") = rqq
    response.cookies("rqz") = rqz
    response.cookies("dw") = dw
    response.cookies("zg") = zg
  

    response.cookies("flag") = "loginok"
    response.cookies("adminuser") = uid

Response.Redirect "qjtycxcl.asp"

%>