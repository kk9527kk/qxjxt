<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>zhanshi</title>

<style type="text/css">
<!--
@import url("../yygs/Css0.Css");
-->
</style>
<link href="css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-top: 0px;
}
-->
</style></head>
<marquee direction="up" scrollamount="1" scrolldelay="120">
<table border="0" cellpadding="0" cellspacing="0">



<% 
hhh =  FormatDateTime(now, vbShortDate)      
sql="select *  from qxjdjb  where (QJZT  >= '5' or QJZT ='10')  and  (ks='1'or ks is null) and (jld='1'or jld is null) and (ld='1'or ld is null)   and XJR is null and QJSJQ <=  #"&hhh&"# and QJSJZ >= #"&hhh&"#   order by id desc"
set rs=server.createobject("adodb.recordset")
//response.write sql
rs.open sql,connstr,1,1
if rs.eof then
	response.write "今天没有人请假"
	response.end
end if

do while not rs.eof
%>


<tr class="top">
<td><%=rs("ZG")%>&nbsp;</td>
<td><%=rs("qjlb")%></td>
<td><%=rs("QJSJQ")%>&nbsp;</td>
<td><%=rs("QJSJZ")%></td>
</tr>

<%
rs.movenext
loop
%>
</table></marquee>

</html>