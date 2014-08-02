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
sql="select * from zg  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw1=rs1("dw")
    zg1=rs1("zg")
    zw1=rs1("zw")
    
end if
rs1.close
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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局请销假查询</title>
</head>

<body background="images/header_1.gif">
<p align="center"><font color="#0000FF" face="黑体" size="6">云阳县国家税务局请销假查询<br>
</font><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></p>

<table border="1" align="center" cellpadding="0" cellspacing="0" width="64%" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="46">
  <tr>
    <td width="100%" height="30" background="images/bg1.gif">
      <p align="center">根据申请时间查询本人请销假情况</td>
  </tr>
  <tr>
    <td width="100%" height="12">
    <form method="POST" action="grcxcl.asp"  name="form1">
           <p align="center">　 </p>
           <p align="center">查询时间起：  <input name="T1" type="text" onClick="WdatePicker()">&nbsp;&nbsp;        
           查询时间止：  <input name="T2" type="text" onClick="WdatePicker()"> </p>         
        <script language="javascript" type="text/javascript" src="rq/My97DatePicker/WdatePicker.js"></script>
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
</html>
