<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
'response.write(uid)

lid=request.QueryString("id")
'response.write(lid)
if lid="" or  request.Cookies("adminuser")="" then
   response.write "��Դδ֪�����ݶ�ʧ!"
   response.end
end if

ksyj=request.form("ksyj")
ks=request.form("KS")

if ksyj="" and ks="1"  then
   ksyj="ͬ���������,�ύ����"
elseif ksyj="" and ks="0" then
   ksyj="��ͬ���������"
else
   ksyj=ksyj
end if
response.write(lid)

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
rs("ksyj")=ksyj
rs("ks")=KS
RS("KSQM")=SQR1
rs("DWrq")=date()
if ks="1" then
	if rs("QJRZW") = "1"  and  rs("ts") <= 3  and rs("SFLFJ") = "NO" then
		rs("QJzt")="10"
		rs.update
		response.write "<script language=JavaScript>{window.alert('�����Ѿ�����!');window.location.href='Sp.asp';}</script>"
		response.end	
	end if
	rs("QJzt")=3
	rs.update
	response.write "<script language=JavaScript>{window.alert('���������Ѿ��ɹ��ύ�����쵼!');window.location.href='Sp.asp';}</script>"
	response.end
 else
   rs("qjzt")="9"
   rs.update
   response.write "<script language=JavaScript>{window.alert('���벻ͬ��,������!');window.location.href='Sp.asp';}</script>"
   response.end
 
end if
end if
rs.close

%>
