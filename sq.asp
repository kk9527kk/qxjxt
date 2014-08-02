<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>

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
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='index.htm';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
sql="select * from zg  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
end if
rs1.close

%>


<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机耗材及配件申请审批表</title>
</head>

<body background="images/header_1.gif" bgproperties="fixed">


<p align="center">                          
　</p>
<p align="center">                          
<form method="POST" action="save.asp?id=<%=uid%>" name="form1">
<table align="center" border="1" width="750" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#FFFFFF" height="232">
  <tr>
    <td width="811" bordercolorlight="#000000" height="36" background="images/bg1.gif" colspan="2">
      <p align="center"><font color="#0000FF" face="黑体" size="6">
      云阳县国家税务局职工请销假审批表</font>
     
    </td>
  </tr>
  <tr>
    <td width="160" bordercolorlight="#000000" height="36">
      <p align="center">                          
申请人：<%=zg%> 
     
    </td>
    <td width="751" bordercolorlight="#000000" height="36">
    请假时间起：&nbsp; <input name="qjsjq" type="text" onClick="WdatePicker()" size="10">&nbsp;&nbsp;&nbsp;                     
    请假时间止：  <input name="qjsjz" type="text" onClick="WdatePicker()" size="12">                        
        <script language="javascript" type="text/javascript" src="rq/My97DatePicker/WdatePicker.js"></script>                    

     
    &nbsp;请假天<input type="text" name="Ts" size="4">天                   
     
    </td>
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">
     
    请假事由,半天假在此备注<textarea rows="2" name="qjsx" cols="90"></textarea>
     
    </td>
  </tr>
  <tr>
    <td width="268" bordercolorlight="#000000" height="36">
     
    请假类别<select size="1" name="qjlb">
      <option value="年假">年&nbsp; 假</option>
      <option value="病假">病&nbsp; 假</option>
      <option value="事假">事&nbsp; 假</option>
      <option value="探亲">探&nbsp; 亲</option>
      <option value="产假">产&nbsp; 假</option>
      <option value="婚假">婚&nbsp; 假</option>
      <option value="丧假">丧&nbsp; 假</option>
	  <option value="加班">加&nbsp; 班</option>
      <option value="补休">补&nbsp; 休</option>

    </select>
     
    </td>
    <td width="539" bordercolorlight="#000000" height="36">
     
前往何地<input type="text" name="qwhd" size="36">&nbsp; 需局长终审标志<input type="checkbox" name="SFLFJ" value="YES">（
详见操作手册）    </td>  
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;请假日期：<%=date()%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
     
    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      <p align="center">　<input type="submit" name="B1" value="提交审批">&nbsp;                                                                    
      <input type="reset" value="全部重填" name="B2"></p>
      　
    </td>
  </tr>
</table>
 </form>
</body>

</html>