<!--#include file="conn.asp"-->
<% dim broom,roomname
   broom=request("broom")
   blink=request("blink")
   roomname=replace(request("roomname")," ","")
   set rs=server.createobject("adodb.recordset")
   sql="select * from deeptree where parentid="&broom&" and content='"&roomname&"'"
   rs.open sql,conn,1,3
   if not rs.eof and not rs.bof then
%>
<html>
<head>
<meta http-equiv='refresh' content='1;url=admin_type.asp'>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<p>&nbsp;</p>
<center><font color=red>�ð���Ѿ������ڴ˴�����,���������...</font>
<p>&nbsp;</p></body></html>
<%
   else
   rs.addnew
   rs("parentid")=broom
   rs("content")=roomname
   rs("link")=blink
   rs("level")=3
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
<center><font color=red>��ӷ���(<%=roomname%>)�ɹ�!...</font>
<p>&nbsp;</p></body></html>
<%end if%>