<!--#include file="conn.asp"-->
<%
lid=request.QueryString("id")
dw=request.form("dw")
'xm=request.form("xm")
zcmc=request.form("zcmc")
xh=request.form("xh")
xlh=request.form("xlh")
pzqk=request.form("pzqk")
je=request.form("je")
zg=request.form("zg")

select case dw
case "�ֳ���"
dwbm="10"
case "�칫��"
dwbm="11"
case "�����"
dwbm="12"
case "˰����"
dwbm="13"
case "���ܿ�"
dwbm="14"
case  "�Ʋƿ�"
dwbm="15"
case "�˽̿�"
dwbm="16"
case "�����"
dwbm="17"
case "��Ϣ����"
dwbm="18"
case "��˰��"
dwbm="19"
case "�����"
dwbm="20"
case  "����һ��"
dwbm="21"
case  "��������"
dwbm="22"
case  "������"
dwbm="23"
case  "ƽ����"
dwbm="24"
case  "��԰��"
dwbm="25"
case  "������"
dwbm="26"
case  "��¡��"
dwbm="27"
case  "������"
dwbm="28"
case  "����˰��"
dwbm="29"

end select


if xh="" then
  xh="����"
end if
if xlh="" then
  xlh="����"
end if
if pzqk="" then
  pzqk="����"
end if



if zg<>"" then
sql="select * from KP where id="&lid
response.write(sql)
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then
rs("dw")=dw
rs("dwbm")=dwbm
rs("zg")=zg
rs("zcmc")=zcmc
rs("xh")=xh
rs("xlh")=xlh
rs("pzqk")=pzqk
rs("je")=je
rs("djrq")=date()
rs("sl")=1
rs.update
response.redirect "index.asp"
end if
end if
%>
