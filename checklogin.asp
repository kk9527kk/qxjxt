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
alert( "��������������ʺţ�"  );
location.href = "javascript:history.back()"  
</script>
<%end if
if userpwd="" then%>
<script language=javascript>  
alert( "����������������룡"  );
location.href = "javascript:history.back()" 
</script>
<%end if%>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From users where username='" &username&"'", conn, 3,3
%>
<%if rs.bof then %>
<script language=javascript>  
alert( "���󣺴��û��������ڣ�"  );
location.href = "javascript:history.back()"  
</script>
<%elseif md5(userpwd)<>rs("userpwd") then%>
<script language=javascript>  
alert("�����������벻��ȷ��");  
location.href = "javascript:history.back()"  
</script>
<%elseif rs("username")<>"��ά��" Then%>
<script language=javascript>  
alert("�㲻�߱���ϵͳ�Ĺ���Ȩ�ޣ�����ϵϵͳ����Ա��");  
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
<center><font color=red>��¼������ɹ�!...</font>
<p>&nbsp;</p></body></html>
<%end if%>