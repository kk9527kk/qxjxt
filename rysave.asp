<!--#include file="conn.asp"-->
<%
dw=request.form("dw")
zg=request.form("zg")
zw=request.form("zw")
lid=request.form("id")
xh=request.form("xh")
if zg="" or lid="" or XH="" OR dw="" then
   response.write "<script language=JavaScript>{window.alert('�˺š���λ����������ž�����Ϊ�գ�');window.location.href='RYADD.asp';}</script>"
   response.end
 else

sql="select * from zg"
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
rs.addnew
rs("dw")=dw
rs("zg")=zg
rs("mm")="111111"
rs("zw")=zw
RS("ID")=lID
rs("xh")=xh
rs.update
response.redirect "zhgl.asp"
END IF
%>
