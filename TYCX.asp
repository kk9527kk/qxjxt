<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
'lid=request.QueryString("id")
uid=request.Cookies("adminuser")
'response.write(uid)
sql="select * from djb  where id='"&uid&"'"
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰��������ٹ���ϵͳ</title>
</head>

<body background="images/header_1.gif">
<p align="center"><font face="����" size="6" color="#0000FF">�����ع���˰�������ٲ�ѯ<br>
</font><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></p>

<table border="1" align="center" cellpadding="0" cellspacing="0" width="64%" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="46">
  <tr>
    <td width="100%" height="30" background="images/bg1.gif">
      <p align="center">���ݸ������������в�ѯ</td>
  </tr>
  <tr>
    <td width="100%" height="12">
    <form method="POST" action="cxcl.asp"  name="form1">
           <p align="center">�� </p>
           <p align="center">��ѯʱ����  
           <input name="T1" type="text" onClick="WdatePicker()" size="20">&nbsp;&nbsp;   
           ��ѯʱ��ֹ��  
           <input name="T2" type="text" onClick="WdatePicker()" size="20"> </p>   
        <script language="javascript" type="text/javascript" src="rq/My97DatePicker/WdatePicker.js"></script>
           <p align="center">������λ�� <select size="1" name="D1">  
           </select>&nbsp;&nbsp; ְ�������� <select size="1" name="D1">  
           </select> </p>
           ��<p align="center"><input type="submit" value="ȷ��" name="B1"><input type="reset" value="��д" name="B2"></p></td>  </form>
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