<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->

<%
dim username,userpwd,rs,sql,userip
userip=Request.ServerVariables("REMOTE_ADDR")
username=trim(replace(request("username"),"'",""))
userpwd=trim(Request.Form("userpwd"))
if username="" then
%>
<script language=javascript>  
alert( "错误：请输入管理帐号！"  );
location.href = "javascript:history.back()"  
</script>
<%end if
if userpwd="" then%>
<script language=javascript>  
alert( "错误：请输入管理密码！"  );
location.href = "javascript:history.back()" 
</script>
<%end if%>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From users where username='" &username&"'", conn, 3,3
%>
<%if rs.bof then %>
<script language=javascript>  
alert( "错误：此用户名不存在！"  );
location.href = "javascript:history.back()"  
</script>
<%elseif md5(userpwd)<>rs("userpwd") then%>
<script language=javascript>  
alert("错误：您的密码不正确！");  
location.href = "javascript:history.back()"  
</script>
<%elseif rs("username")<>"沈维寿" Then%>
<script language=javascript>  
alert("你不具备本系统的管理权限，请联系系统管理员！");  
location.href = "javascript:history.back()"  
</script>
<%else%>

<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From users where username='" &username&"'", conn, 3,3
%>
<%
session("username")=rs("username")
rs.close
Set rs=nothing
%>
<html>
<head>
<meta http-equiv='refresh' content='1;url=admin_type.asp'>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<p>&nbsp;</p>
<center><font color=red>登录管理版块成功!...</font>
<p>&nbsp;</p></body></html>
<%end if%>