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
<STYLE type=text/css>A:link {
	COLOR: #000000; FONT-FAMILY: ����; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; FONT-FAMILY: ����; TEXT-DECORATION: none
}
A:active {
	FONT-FAMILY: ����; TEXT-DECORATION: underline
}
A:hover {
	COLOR: #84bd6b; TEXT-DECORATION: underline overline
}
BODY {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
TABLE {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
.f24 {
	COLOR: #ff0000; FONT-FAMILY: ����; FONT-SIZE: 9pt; TEXT-DECORATION: underline overline
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
<title>�����ع���˰��ּ�����Ĳļ��������������</title>
</head>

<body background="images/header_1.gif">
<%sql="select * from sqspb  where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if  rs.eof then
    response.write "<script language=JavaScript>{window.alert('�Բ���,û�ҵ���������Ϣ!');window.location.href='sp.asp';}</script>"
    response.end
else

%>
<p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰��ּ�����Ĳļ��������������</font></p>
<p align="center"><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
<table align="center" border="1" width="75%" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolor="#000000" height="261">
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="36" background="images/bg1.gif">
     
<p align="center">                          
���뵥λ:<%=rs("dw")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                                   
������:<%=rs("sqr")%></p>
     
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="36">����Ĳ���:(<font color="#008080"><%=rs("hcmc")%></font>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;          
      �Ĳ��ͺ�:(<font color="#008080"><%=rs("xh")%></font>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;              
      ����:(<font color="#008080"><%=rs("sl")%></font>)
     
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="45">ʹ����:(<font color="#008080"><%=rs("syr")%></font>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     
      ��������:(<font color="#008080"><%=rs("sqsx")%></font>)</td>
  </tr>
  <form method="POST" action="savezx.asp?id=<%=lid%>" name="form1">
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="15">��λ���������������[<font color="#008080">
<%=rs("ksyj")%></font> ]      <p>&nbsp;ǩ����<%=rs("ks")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                             
      ʱ�䣺<%=rs("ksrq")%></p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="15">��Ϣ����ǩ��:                                                                       
      ͬ��<input type="radio" value="1" name="zx" checked>����ͬ��<input type="radio" name="zx" value="0"><textarea rows="2" name="zxyj" cols="82"></textarea>
      <p>&nbsp;ǩ����<%=zg%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;          
      ʱ�䣺<%=date()%></p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#FFFFFF" height="16">
      <p align="center">��<input type="submit" name="B1" value="�ύ����">&nbsp;        
      <input type="reset" value="ȫ����д" name="B2"></p>
      <p>��</p>
    </td></form>
  </tr>
</table>
 
</body>
<%end if %>
<%end if %></html>
