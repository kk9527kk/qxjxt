<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
lid=request.QueryString("id")

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
  ZG1=RS("ZG")
  SFLFJ1=RS("SFLFJ")
END IF
RS.CLOSE

sql="select * from ZG where ZG='"&ZG1&"'"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,connstr,3,2
if not rs2.eof then
    Zw1=rs2("Zw")
end if
rs2.close


sql="select * from QXJDJb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("RJyj")=ksyj
rs("RJ")=SQR1
RS("RJK")=SQR1
rs("RJKrq")=date()
if ks="0"   then
  rs("QJzt")="9"
   rs.update
   response.write "<script language=JavaScript>{window.alert('���벻ͬ��,������!');window.location.href='SpRJ.asp';}</script>"
   response.end
END IF
IF KS="1" AND SFLFJ1="NO" AND (ZW1="1" or zw1="7") THEN
     rs("QJzt")="6"
     RS("ZGZT")="1"
   rs.update
   response.write "<script language=JavaScript>{window.alert('����ͬ��,������!');window.location.href='SpRJ.asp';}</script>"
   response.end
   END IF 
IF KS="1" AND SFLFJ1="YES" AND (ZW1="1" or zw1="7") THEN
     rs("QJzt")="8"
   rs.update
   response.write "<script language=JavaScript>{window.alert('�����˽̿�����,�ɹ����ֳ�����!');window.location.href='SpRJ.asp';}</script>"
   response.end
   END IF 
IF KS="1" AND SFLFJ1="NO" AND ZW1="2" THEN
     rs("QJzt")="3"
   rs.update
   response.write "<script language=JavaScript>{window.alert('�����˽̿�����,�ɹ����ֳֹܾ�����!');window.location.href='SpRJ.asp';}</script>"
   response.end
   END IF 
IF KS="1" AND SFLFJ1="YES" AND ZW1="2" THEN
     rs("QJzt")="8"
   rs.update
   response.write "<script language=JavaScript>{window.alert('�����˽̿�����,�ɹ����ֳ�����!');window.location.href='SpRJ.asp';}</script>"
   response.end
END IF 

IF KS="1" AND (ZW1="3" OR ZW1="6") THEN
     rs("QJzt")="3"
   rs.update
   response.write "<script language=JavaScript>{window.alert('�����˽̿�����,�ɹ����ֳֹܾ�����!');window.location.href='SpRJ.asp';}</script>"
   response.end
END IF 

IF KS="1" AND ZW1="4"  THEN
     rs("QJzt")="8"
   rs.update
   response.write "<script language=JavaScript>{window.alert('�����˽̿�����,�ɹ����ֳ�����!');window.location.href='SpRJ.asp';}</script>"
   response.end
END IF 
 
end if

%>
