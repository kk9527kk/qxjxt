<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
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
   response.write "<script language=JavaScript>{window.alert('�Բ���,����Ȩ��������!');window.location.href='sp1.asp';}</script>"
   response.end
else
%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰��ּ�����Ĳļ��������������</title>
</head>

<body>
<%sql="select * from sqspb  where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if  rs.eof then
    response.write "<script language=JavaScript>{window.alert('�Բ���,û�ҵ���������Ϣ!');window.location.href='sp.asp';}</script>"
    response.end
else

%>
<p align="center"><font size="5" face="����_GB2312"><b>�����ع���˰��ּ�����Ĳļ��������������</b></font></p>
<p align="center">��</p>
<p align="center">                          
���뵥λ:<%=rs("dw")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                   
������:<%=rs("sqr")%></p>
<table align="center" border="1" width="68%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#000000" height="261">
  <tr>
    <td width="25%" bordercolorlight="#FFFFFF" height="36">����Ĳ���:(<font color="#FF00FF"><%=rs("hcmc")%></font>)
     
    </td>
    <td width="36%" bordercolorlight="#FFFFFF" height="36">�Ĳ��ͺ�:(<font color="#FF00FF"><%=rs("xh")%></font>)</td>
    <td width="24%" bordercolorlight="#FFFFFF" height="36">����:(<font color="#FF00FF"><%=rs("sl")%></font>)</td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="45" colspan="3">��������:(<font color="#FF00FF"><%=rs("sqsx")%></font>)</td>
  </tr>
  <form method="POST" action="savezx.asp?id=<%=lid%>" name="form1">
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="15" colspan="3">��λ������ǩ��:                                                       
      ͬ��<input type="radio" value="1" name="DWFZR" checked>����ͬ��<input type="radio" name="DWFZR" value="0"><textarea rows="2" name="ksyj" cols="65"><%=rs("ksyj")%></textarea>
      <p>&nbsp;ǩ����<%=rs("ks")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;             
      ʱ�䣺<%=rs("ksrq")%></p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="15" colspan="3">��Ϣ����ǩ��:                                                       
      ͬ��<input type="radio" value="1" name="zx" checked>����ͬ��<input type="radio" name="zx" value="0"><textarea rows="2" name="zxyj" cols="67"></textarea>
      <p>&nbsp;ǩ����<%=zg%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      ʱ�䣺<%=date()%></p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="16" colspan="3">
      <p align="center">��<input type="submit" name="B1" value="�ύ����">&nbsp; 
      <input type="reset" value="ȫ����д" name="B2"></p>
      ��
      <p>��</p>
    </td>
  </tr>
  <tr>
    <td width="25%" bordercolorlight="#FFFFFF" height="16">��</td>
    <td width="36%" bordercolorlight="#FFFFFF" height="16">��</td>
    <td width="24%" bordercolorlight="#FFFFFF" height="16">��</td>
  </tr>
</table>
 </form>
</body>
<%end if %>
<%end if %></html>
