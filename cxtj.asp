<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
'lid=request.QueryString("id")
uid=request.Cookies("adminuser")
'response.write(uid)
sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("zw")
end if
rs1.close
if zw<>"3" and zw<>"4"  and zw<>"6" then
   response.write "<script language=JavaScript>{window.alert('对不起,你无权进行查询!');window.location.href='sp1.asp';}</script>"
   response.end
else
%>


<html>

<head>
<STYLE type=text/css>A:link {
	COLOR: #000000; FONT-FAMILY: 宋体; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; FONT-FAMILY: 宋体; TEXT-DECORATION: none
}
A:active {
	FONT-FAMILY: 宋体; TEXT-DECORATION: underline
}
A:hover {
	COLOR: #84bd6b; TEXT-DECORATION: underline overline
}
BODY {
	COLOR: #000000; FONT-FAMILY: 宋体; FONT-SIZE: 9pt
}
TABLE {
	COLOR: #000000; FONT-FAMILY: 宋体; FONT-SIZE: 9pt
}
.f24 {
	COLOR: #ff0000; FONT-FAMILY: 宋体; FONT-SIZE: 9pt; TEXT-DECORATION: underline overline
}
.l17 {
	LINE-HEIGHT: 170%
}
.f18 {
	FONT-SIZE: 18px
}
</STYLE>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机耗材及配件申请审批表</title>
</head>

<body background="images/header_1.gif">
<p align="center"><font color="#0000FF" face="黑体" size="6">云阳县国家税务局计算机耗材及配件统计<br>
</font><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></p>

<table border="1" align="center" cellpadding="0" cellspacing="0" width="64%" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="46">
  <tr>
    <td width="100%" height="30" background="images/bg1.gif">
      <p align="center">按时间段分单位进行统计</td>
  </tr>
  <tr>
    <td width="100%" height="12">
    <form method="POST" action="tjcxcl.asp"  name="form1">
      <p align="center">　</p>
        
    
      <p align="center">请输入查询时间起
      
      
      <script language="javascript" type="text/javascript" src="rq/My97DatePicker/WdatePicker.js"></script>


<input name="T1" type="text" onClick="WdatePicker()" size="20"> </p>
      
      
        
    
      <p align="center">请输入查询时间止(格式：2008/12/31)：<input type="text" name="T2" size="20"></p>
      <p>　<p align="center"><input type="submit" value="确认" name="B1"><input type="reset" value="重写" name="B2"></p></td>  </form>
  </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="64%" height="55">
  <tr>
    <td width="100%" height="55">
    </td>
  </tr>
</table>
</body>
<%end if %></html>