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
sql="select * from ZG  where id='"&uid&"'"
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
<title>云阳县国家税务局计算机耗材及配件申请审批文书</title>
</head>

<body background="images/header_1.gif">
<%sql="select * from QXJDJb  where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if  rs.eof then
    response.write "<script language=JavaScript>{window.alert('对不起,没找到该信息!');window.location.href='XH.asp';}</script>"
    response.end
else

%>
<p align="center"><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>       
<table align="center" border="1" width="750" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#FFFFFF" height="232">
  <tr>
    <td width="811" bordercolorlight="#000000" height="36" background="file:///D:/qxjxt/images/bg1.gif" colspan="2">
      <p align="center"><font color="#0000FF" face="黑体" size="6">云阳县国家税务局职工请销假审批表</font></td>
  </tr>
  <tr>
    <td width="160" bordercolorlight="#000000" height="36">
      <p align="center">申请人：<%=rs("zg")%>
    </td>
    <td width="751" bordercolorlight="#000000" height="36">请假时间起 ： <font color="#FF00FF">&nbsp;<%=rs("qjsjq")%>           
      </font>&nbsp;&nbsp;&nbsp; 请假时间止： <font color="#FF00FF">&nbsp; <%=rs("qjsjz")%>           
      </font> &nbsp;共请假<font color="#FF00FF">&nbsp;<%=rs("ts")%>
      </font><font color="#000000">天</font></td>       
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">请假事由：(<%=rs("qjsx")%>
      )</td>       
  </tr>
  <tr>
    <td width="268" bordercolorlight="#000000" height="36">请假类别：(<font color="#FF00FF">&nbsp;<%=rs("qjLB")%>
      </font>)</td>
    <% 
   SFLFJ=RS("SFLFJ")
   IF SFLFJ="YES" THEN
     SFLFJ1="是"
   ELSE 
     SFLFJ1="否"
   END IF  
%>
<td width="539" bordercolorlight="#000000" height="36">前往何地：(<font color="#FF00FF">&nbsp;<%=rs("qWHD")%>
      </font>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 需局长终审标志：(<font color="#FF00FF">&nbsp;<%=SFLFJ1%>       
      </font>)</td>
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;请假日期：<font color="#FF00FF">&nbsp;<%=RS("SQSJ")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>       
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">单位领导审批意见：<font color="#FF00FF">&nbsp;(<%=rs("ksyj")%>
      )</font>       
      <p>&nbsp;签名：<font color="#FF00FF">&nbsp;<%=RS("KSQM")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        
      时间：<font color="#FF00FF">&nbsp;<%=RS("DWRQ")%>
      </font></td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">年假、探亲、婚（丧）假政策规定、人教科意见：        
      (<font color="#FF00FF">&nbsp;<%=RS("RJYJ")%>
      </font>)
      <p>&nbsp;签名：<font color="#FF00FF">&nbsp;<%=RS("RJ")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      签字时间：<font color="#FF00FF">&nbsp;<%=rs("rjkrq")%>
      </font></td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">分管局领导审批意见：        
      (<font color="#FF00FF">&nbsp;<%=rs("fgldyj")%>
      </font>)<p>&nbsp;签名：<font color="#FF00FF">&nbsp;<%=rs("fgld")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      签审时间：<font color="#FF00FF">&nbsp;<%=rs("fgldrq")%>
      </font></td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">县局局长意见：(<font color="#FF00FF">&nbsp;<%=rs("JZYJ")%></font>) 
      <p>&nbsp;签名：<font color="#FF00FF">&nbsp;<%=rs("JZ")%>
      </font>      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      签审时间：<font color="#FF00FF">&nbsp;<%=rs("JZrq")%>
      </font>    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">请销假情况：实休天数(<font color="#FF00FF">&nbsp;<%=rs("sxts")%>
      </font>)<p>&nbsp;销假人：<font color="#FF00FF">&nbsp;<%=rs("XJR")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      销假日期：<font color="#FF00FF">&nbsp;<%=rs("XJRQ")%>
      </font>&nbsp;    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      <p align="center">　&nbsp; <a href="javascript:history.go(-1);">返回</a>  
     
      </p>
      </form>　</td>
  </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="66%" height="55">
  <tr>
    
  </tr>
</table>
</body>
<%end if %>
</html>





