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
uid=request.Cookies("adminuser")
lid=request.QueryString("id")
if lid="" or  request.Cookies("adminuser")="" then
   response.write "��Դδ֪�����ݶ�ʧ!"
   response.end
end if
qssl=request.form("qssl")
je=request.form("je")

if trim(qssl)="" OR  trim(je)="" then
   response.write "<script language=JavaScript>{window.alert('ʵ�������ͽ���Ϊ�� ��');window.location.href='cxXH.asp';}</script>"
RESPONSE.END
end if
    
  'isInteger  IsNumeric
  if not(isInteger(qssl) and IsNumeric(je)) THEN
     response.write "<script language=JavaScript>{window.alert(' ����Ϊ��������������Ϊ��ʵ�� ��');window.location.href='cxXH.asp';}</script>"
  RESPONSE.END
END IF
  if qssl<=0 then
   response.write "<script language=JavaScript>{window.alert(' ����Ϊ��������������Ϊ��ʵ�� ��');window.location.href='cxXH.asp';}</script>"
  RESPONSE.END
END IF  
   if IsNumeric(je) THEN
       if je<=0 then
     response.write "<script language=JavaScript>{window.alert(' ����Ϊ��������������Ϊ��ʵ�� ��');window.location.href='cxXH.asp';}</script>"
  RESPONSE.END
  end if
END IF




sql="select * from sqspb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("qssl")=qssl
rs("je")=je
rs("xhrq")=date()
rs("xhbz")="1"
   rs.update
   response.write "<script language=JavaScript>{window.alert('�������ųɹ�!');window.location.href='cxxh.asp';}</script>"
   response.end
end if
 


%>
