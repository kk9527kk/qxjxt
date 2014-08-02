<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
lid=request.QueryString("id")
uid=request.Cookies("adminuser")
sql="select * from ZG  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("zw")
end if
rs1.close

if zw="1" OR zw="4" then
   response.write "<script language=JavaScript>{window.alert('对不起,你无权进行审批!');window.location.href='sp.asp';}</script>"
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
<title>云阳县国家税务局职工请销假审批表</title>
</head>

<body background="images/header_1.gif" bgproperties="fixed">
<%sql="select * from QXJDJB  where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if  rs.eof then
    response.write "<script language=JavaScript>{window.alert('对不起,没找到该审批信息!');window.location.href='sp.asp';}</script>"
    response.end
else

%>
<table border="0" id="table1" align="center" cellspacing="0" cellpadding="0" width="691" height="429">
  <tr>
<td>
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
申请人：<%=rs("zg")%> 
     
    </td>
    <td width="751" bordercolorlight="#000000" height="36">请假时间起 ：                          
    <font color="#FF00FF">&nbsp;<%=rs("qjsjq")%></font>&nbsp;&nbsp;&nbsp;                                                      
    请假时间止： <font color="#FF00FF">&nbsp; <%=rs("qjsjz")%></FONT>    &nbsp;共请假<font color="#FF00FF">&nbsp;<%=rs("ts")%></FONT><font color="#000000">天</font>                                                   
     
    </td>
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">
     
    请假事由：(<%=rs("qjsx")%>)
     
    </td>
  </tr>
  <tr>
    <td width="268" bordercolorlight="#000000" height="36">
     
    请假类别：(<font color="#FF00FF">&nbsp;<%=rs("qjLB")%></FONT>)     
    </td>
    <% 
   SFLFJ=RS("SFLFJ")
   IF SFLFJ="YES" THEN
     SFLFJ1="是"
   ELSE 
     SFLFJ1="否"
   END IF  
%>
    <td width="539" bordercolorlight="#000000" height="36">
     
前往何地：(<font color="#FF00FF">&nbsp;<%=rs("qWHD")%></FONT>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
需局长审批标志：(<font color="#FF00FF">&nbsp;<%=SFLFJ1%></FONT>)    </td>                                  
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;请假日期：<font color="#FF00FF">&nbsp;<%=RS("SQSJ")%></FONT>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
     
    </td>
  </tr>
  <tr>
      
  <form method="POST" action="saveRJK.asp?id=<%=lid%>" name="form1">
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      单位负责人意见：<font color="#FF00FF">&nbsp;(<%=rs("ksyj")%>)</FONT>                 
      <p>&nbsp;签名：<%=RS("KSQM")%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;              
      时间：<font color="#FF00FF">&nbsp;<%=RS("DWRQ")%></FONT>
    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      人教科意见： 同意<input type="radio" value="1" name="KS" checked>，不同意<input type="radio" name="KS" value="0">&nbsp;&nbsp;           
      在此填上具体意见<textarea rows="2" name="KSyj" cols="59"></textarea>                  
      <p>&nbsp;签名：<%=zg%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 时间：<%=DATe()%> 
    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      <p align="center">　<input type="submit" name="B1" value="提交审批">&nbsp;                                                                        
      <input type="reset" value="全部重填" name="B2"></p>
     </form> 　
    </td>
  </tr>
</table>
 


    <td width="689" height="18">
    </td>
  </tr>
</table>
</body>
<%end if %>
<%end if
rs.close
set rs=nothing %>
</html>