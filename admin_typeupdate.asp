<!--#include file="conn.asp"-->
<%    
   set rs=server.createobject("adodb.recordset")
   sql="select * from deeptree where id="&request("id")
   rs.open sql,conn,1,3
   rs("parentid")=request("broom1")
   rs("content")=request("typename")
   rs("link")=request("blink")
   rs.update
%>
<html>
<head>
<meta http-equiv='refresh' content='1;url=admin_type.asp'>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<p>&nbsp;</p>
<center><font color=red>���İ����Ϣ�ɹ�!...</font>
<p>&nbsp;</p></body></html>