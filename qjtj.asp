<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
lid=request.QueryString("id")
uid=request.Cookies("adminuser")
'response.write(uid)
sql="select * from qxjdjb  where zg='"&uid&"'"
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
<title>云阳县国家税务局请销假管理系统</title>
</head>

<body background="images/header_1.gif">
<p align="center"><font color="#0000FF" face="黑体" size="6">云阳县国家税务局请销假查询<br>
</font><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></p>

<table border="1" align="center" cellpadding="0" cellspacing="0" width="64%" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="46">
  <tr>
    <td width="100%" height="30" background="images/bg1.gif">
      <p align="center">根据给定的条件进行查询统计</td>
  </tr>
  <tr>
    <td width="100%" height="12">
    <form method="POST" action="qjtjb.asp"  name="form1">
           <p align="center">　 </p>
           <p align="center">查询时间起：  <input name="T1" type="text" onClick="WdatePicker()" size="10">&nbsp;&nbsp;     
           查询时间止：  <input name="T2" type="text" onClick="WdatePicker()" size="11">           
           &nbsp; </p>                             
           <p align="center">&nbsp;查询单位： <select size="1" name="dw">       
            <option value="">全局</option>
            
            
            
<% sql200="select *  from dw order by dwbm  "
set rs200=server.createobject("adodb.recordset")
rs200.open sql200,connstr,1,1           
do while not rs200.eof
%>
<option value="<%=rs200("dw")%>"><%=rs200("dw")%></option>
<%rs200.movenext
loop    
rs200.close
%>      
           </select>&nbsp;&nbsp;&nbsp; 休假类别：<select size="1" name="qjlb">  
            <option  value="">全部</option>
      <option value="年假">年&nbsp; 假</option>
      <option value="病假">病&nbsp; 假</option>
      <option value="事假">事&nbsp; 假</option>
      <option value="探亲">探&nbsp; 亲</option>
      <option value="产假">产&nbsp; 假</option>
      <option value="婚假">婚&nbsp; 假</option>
      <option value="丧假">丧&nbsp; 假</option>
	  <option value="加班">加&nbsp; 班</option>
      <option value="补休">补&nbsp; 休</option>
           </select> </p>                     
        <script language="javascript" type="text/javascript" src="rq/My97DatePicker/WdatePicker.js"></script>
           <p>　<p align="center"><input type="submit" value="确认" name="B1"><input type="reset" value="重写" name="B2"></p></td>  </form>
  </tr>
</table>

</body>


