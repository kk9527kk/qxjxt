<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%function isInteger(para) 
on error resume next 
dim str 
dim l,i 
if isNUll(para) then 
isInteger=false 
exit function 
end if 
str=cstr(para) 
if trim(str)="" then 
isInteger=false 
exit function 
end if 
l=len(str) 
for i=1 to l 
if mid(str,i,1)>"9" or mid(str,i,1)<"0" then 
isInteger=false 
exit function 
end if 
next 
isInteger=true 
if err.number<>0 then err.clear 
end function %>


<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
lid=request.QueryString("id")
if lid="" or lid<>request.Cookies("adminuser") then
response.write "��Դδ֪�����ݶ�ʧ!"
response.end
end if
qjsjq=request.form("qjsjq")
qjsjz=request.form("qjsjz")
ts=request.form("ts")
qjlb=request.form("qjlb")
qjsx=request.form("qjsx")
qjlb=request.form("qjlb")
QWHD=REQUEST.FORM("QWHD")
SFLFJ=REQUEST.FORM("SFLFJ")
IF SFLFJ="YES" THEN
  SFLFJ="YES"
 ELSE
  SFLFJ="NO"
 END IF
 if DateDiff("d",qjsjq,qjsjz)<0 then 
   response.write "<script language=JavaScript>{window.alert('���ֹʱ��Ӧ���������ʱ�䣬�������룡');window.location.href='sq.asp';}</script>"
   response.end
end if

if qjsjq="" then
  response.write "<script language=JavaScript>{window.alert('���ʱ��������Ϊ�գ�');window.location.href='SQ.asp';}</script>"
response.end
end if
if qjsjz="" then
  response.write "<script language=JavaScript>{window.alert('���ʱ��ֹ������Ϊ�գ�');window.location.href='SQ.asp';}</script>"
response.end
end if

if qjsx="" then
  response.write "<script language=JavaScript>{window.alert('����������Ϊ�գ�');window.location.href='SQ.asp';}</script>"
response.end
end if

if qjsX="" then
  response.write "<script language=JavaScript>{window.alert('�������Ϊ�գ�');window.location.href='SQ.asp';}</script>"
response.end
end if

if (not isInteger(TS)) then
   response.write "<script language=JavaScript>{window.alert('��������Ϊ��Ȼ����');window.location.href='SQ.asp';}</script>"
response.end
end if

if QWHD="" then
  response.write "<script language=JavaScript>{window.alert('ǰ���εز�����Ϊ�գ�');window.location.href='SQ.asp';}</script>"
response.end
end if


'sqr=request.cookies("adminuser")

sql="select * from zg  where id='"&lid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg1=rs1("zg")
    zw=rs1("zw")
 '   fgld=rs1("fgld")
end if
sxts=0
select case zw
  case "1"
    qjzt="1"
    message="�����Ѿ��ɹ��ύ,�ȴ������쵼����"
  case  "2"
    qjzt="1"
     message="�����Ѿ��ɹ��ύ,�ȴ������쵼����"
  case "3"
    qjzt="3"
    message="�����Ѿ��ɹ��ύ,�ȴ����쵼����"
  case  "4"
    qjzt="4"
    message="�����Ѿ��ɹ��ύ,�ȴ��ֳ�����"
  case "5"
    qjzt="5"
    sxts=ts
    message="�����Ѿ��ɹ����棬��������"
  case  "6"
    qjzt="2"
    message="�����Ѿ��ɹ��ύ,�ȴ��˽̿�����"
  case "7" 
    qjzt="1"
    message="�����Ѿ��ɹ��ύ,�ȴ�����������"
 end select
rs1.close
sql="select * from qxjdjb"
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
rs.addnew
rs("qjsx")=qjsx
rs("qjsjq")=qjsjq
rs("qjsjz")=qjsjz
rs("qjlb")=qjlb
rs("ts")=ts
rs("qwhd")=qwhd
rs("sqsj")=date()
rs("sxts")=sxts
rs("dw")=dw
rs("zg")=zg1
rs("qjzt")=qjzt
rs("zgzt")="2"
rs("SFLFJ")=SfLFJ
rs("QJRZW")=zw
rs.update
response.write "<script language=JavaScript>{window.alert('"&message&"');window.location.href='SQ.asp';}</script>"
    response.end
%>
