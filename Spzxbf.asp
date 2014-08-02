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
sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("zw")
end if
rs1.close
if zw<>"3"  then
   response.write "<script language=JavaScript>{window.alert('对不起,你无权进行审批!');window.location.href='sp1.asp';}</script>"
   response.end
else
%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机耗材及配件申请审批表</title>
</head>

<body>
<%sql="select * from sqspb  where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if  rs.eof then
    response.write "<script language=JavaScript>{window.alert('对不起,没找到该审批信息!');window.location.href='sp.asp';}</script>"
    response.end
else

%>
<p align="center"><font size="5" face="楷体_GB2312"><b>云阳县国家税务局计算机耗材及配件申请审批表</b></font></p>
<p align="center">　</p>
<p align="center">                          
申请单位:<%=rs("dw")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                   
申请人:<%=rs("sqr")%></p>
<table align="center" border="1" width="68%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#000000" height="261">
  <tr>
    <td width="25%" bordercolorlight="#FFFFFF" height="36">所需耗材名:(<font color="#FF00FF"><%=rs("hcmc")%></font>)
     
    </td>
    <td width="36%" bordercolorlight="#FFFFFF" height="36">耗材型号:(<font color="#FF00FF"><%=rs("xh")%></font>)</td>
    <td width="24%" bordercolorlight="#FFFFFF" height="36">数量:(<font color="#FF00FF"><%=rs("sl")%></font>)</td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="45" colspan="3">申请事项:(<font color="#FF00FF"><%=rs("sqsx")%></font>)</td>
  </tr>
  <form method="POST" action="savezx.asp?id=<%=lid%>" name="form1">
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="15" colspan="3">单位负责人签字:                                                       
      同意<input type="radio" value="1" name="DWFZR" checked>，不同意<input type="radio" name="DWFZR" value="0"><textarea rows="2" name="ksyj" cols="65"><%=rs("ksyj")%></textarea>
      <p>&nbsp;签名：<%=rs("ks")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;             
      时间：<%=rs("ksrq")%></p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="15" colspan="3">信息中心签字:                                                       
      同意<input type="radio" value="1" name="zx" checked>，不同意<input type="radio" name="zx" value="0"><textarea rows="2" name="zxyj" cols="67"></textarea>
      <p>&nbsp;签名：<%=zg%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      时间：<%=date()%></p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="16" colspan="3">
      <p align="center">　<input type="submit" name="B1" value="提交审批">&nbsp; 
      <input type="reset" value="全部重写" name="B2"></p>
      　
      <p>　</p>
    </td>
  </tr>
  <tr>
    <td width="25%" bordercolorlight="#FFFFFF" height="16">　</td>
    <td width="36%" bordercolorlight="#FFFFFF" height="16">　</td>
    <td width="24%" bordercolorlight="#FFFFFF" height="16">　</td>
  </tr>
</table>
 </form>
</body>
<%end if %>
<%end if %></html>
