<!--#include file="conn.asp"-->
<%
   set rs=server.createobject("adodb.recordset")

   'sql="delete from reply where titleid in(select titleid from article where roomid in(select id from deeptree where parentid="&request("id")&"))"
   'rs.open sql,conn,1,1

   'sql="delete from reply where titleid in(select titleid from article where roomid="&request("id")&")"
   'rs.open sql,conn,1,1

   'sql="delete from article where roomid in(select id from deeptree where parentid="&request("id")&")"
   'rs.open sql,conn,1,1

   'sql="delete from article where roomid="&request("id")
   'rs.open sql,conn,1,1

   sql="delete from deeptree where id in(select id from deeptree where parentid="&request("id")&")"
   rs.open sql,conn,1,1

   sql="delete from deeptree where id="&request("id")
   rs.open sql,conn,1,1
 
%>
<html>
<head>
<meta http-equiv='refresh' content='1;url=admin_type.asp'>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<p>&nbsp;</p>
<center><font color=red>删除版块成功!...</font>
<p>&nbsp;</p></body></html>