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
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='index.htm';}</script>"
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
<title>�����ع���˰��ּ�����Ĳļ��������������</title>
</head>

<body background="images/header_1.gif" bgproperties="fixed">


<p align="center">                          
��</p>
<p align="center">                          
<form method="POST" action="save.asp?id=<%=uid%>" name="form1">
<table align="center" border="1" width="750" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#FFFFFF" height="232">
  <tr>
    <td width="811" bordercolorlight="#000000" height="36" background="images/bg1.gif" colspan="2">
      <p align="center"><font color="#0000FF" face="����" size="6">
      �����ع���˰���ְ��������������</font>
     
    </td>
  </tr>
  <tr>
    <td width="160" bordercolorlight="#000000" height="36">
      <p align="center">                          
�����ˣ�<%=zg%> 
     
    </td>
    <td width="751" bordercolorlight="#000000" height="36">
    ���ʱ����&nbsp; <input name="qjsjq" type="text" onClick="WdatePicker()" size="10">&nbsp;&nbsp;&nbsp;                     
    ���ʱ��ֹ��  <input name="qjsjz" type="text" onClick="WdatePicker()" size="12">                        
        <script language="javascript" type="text/javascript" src="rq/My97DatePicker/WdatePicker.js"></script>                    

     
    &nbsp;�����<input type="text" name="Ts" size="4">��                   
     
    </td>
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">
     
    �������,������ڴ˱�ע<textarea rows="2" name="qjsx" cols="90"></textarea>
     
    </td>
  </tr>
  <tr>
    <td width="268" bordercolorlight="#000000" height="36">
     
    ������<select size="1" name="qjlb">
      <option value="���">��&nbsp; ��</option>
      <option value="����">��&nbsp; ��</option>
      <option value="�¼�">��&nbsp; ��</option>
      <option value="̽��">̽&nbsp; ��</option>
      <option value="����">��&nbsp; ��</option>
      <option value="���">��&nbsp; ��</option>
      <option value="ɥ��">ɥ&nbsp; ��</option>
	  <option value="�Ӱ�">��&nbsp; ��</option>
      <option value="����">��&nbsp; ��</option>

    </select>
     
    </td>
    <td width="539" bordercolorlight="#000000" height="36">
     
ǰ���ε�<input type="text" name="qwhd" size="36">&nbsp; ��ֳ������־<input type="checkbox" name="SFLFJ" value="YES">��
��������ֲᣩ    </td>  
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ڣ�<%=date()%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
     
    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      <p align="center">��<input type="submit" name="B1" value="�ύ����">&nbsp;                                                                    
      <input type="reset" value="ȫ������" name="B2"></p>
      ��
    </td>
  </tr>
</table>
 </form>
</body>

</html>