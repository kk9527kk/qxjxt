<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
lid=request.QueryString("id")
uid=request.Cookies("adminuser")
response.write(uid)
sql="select * from ZG  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("zw")
end if
rs1.close

if zw<>"4" then
   response.write "<script language=JavaScript>{window.alert('�Բ���,����Ȩ��������!');window.location.href='sp.asp';}</script>"
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
<title>�����ع���˰���ְ��������������</title>
</head>

<body background="images/header_1.gif" bgproperties="fixed">
<%sql="select * from QXJDJB  where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if  rs.eof then
    response.write "<script language=JavaScript>{window.alert('�Բ���,û�ҵ���������Ϣ!');window.location.href='sp.asp';}</script>"
    response.end
else

%>
<table border="0" id="table1" align="center" cellspacing="0" cellpadding="0" width="691" height="429">
  <tr>
<td>
<table align="center" border="1" width="750" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#FFFFFF" height="232">
  <tr>
    <td width="811" bordercolorlight="#000000" height="36" background="images/bg1.gif" colspan="2">
      <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰���ְ��������������</font>
     
    </td>
  </tr>
  <tr>
    <td width="160" bordercolorlight="#000000" height="36">
      <p align="center">                          
�����ˣ�<%=rs("zg")%> 
     
    </td>
    <td width="751" bordercolorlight="#000000" height="36">���ʱ���� ��                                   
    <font color="#FF00FF">&nbsp;<%=rs("qjsjq")%></font>&nbsp;&nbsp;&nbsp;                                                               
    ���ʱ��ֹ�� <font color="#FF00FF">&nbsp; <%=rs("qjsjz")%></FONT>    &nbsp;�����<font color="#FF00FF">&nbsp;<%=rs("ts")%></FONT><font color="#000000">��</font>                                                            
     
    </td>
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">
     
    ������ɣ�(<%=rs("qjsx")%>)
     
    </td>
  </tr>
  <tr>
    <td width="268" bordercolorlight="#000000" height="36">
     
    ������(<font color="#FF00FF">&nbsp;<%=rs("qjLB")%></FONT>)     
    </td>
    <% 
   SFLFJ=RS("SFLFJ")
   IF SFLFJ="YES" THEN
     SFLFJ1="��"
   ELSE 
     SFLFJ1="��"
   END IF  
%>
    <td width="539" bordercolorlight="#000000" height="36">
     
ǰ���εأ�(<font color="#FF00FF">&nbsp;<%=rs("qWHD")%></FONT>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��ֳ������־��(<font color="#FF00FF">&nbsp;<%=SFLFJ1%></FONT>)    </td>                                           
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ڣ�<font color="#FF00FF">&nbsp;<%=RS("SQSJ")%></FONT>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
     
    </td>
  </tr>
  <tr>
      
  <form method="POST" action="saveJLDK.asp?id=<%=lid%>" name="form1">
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      ��λ�����������<font color="#FF00FF">&nbsp;(<%=rs("ksyj")%>)</FONT>                 
      <p>&nbsp;ǩ����<font color="#FF00FF">&nbsp;<%=RS("KSQM")%></FONT>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                       
      ʱ�䣺<font color="#FF00FF">&nbsp;<%=RS("DWRQ")%></FONT>
    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      �˽̿������ (<font color="#FF00FF">&nbsp;<%=RS("RJYJ")%></FONT>)                                
      <p>&nbsp;ǩ����<font color="#FF00FF">&nbsp;<%=RS("RJ")%></FONT>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ʱ�䣺<font color="#FF00FF">&nbsp;<%=DATe()%>  </FONT>             
    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      �ֹ��쵼����� ͬ��<input type="radio" value="1" name="KS" checked>����ͬ��<input type="radio" name="KS" value="0">&nbsp;&nbsp;     
      �ڴ����Ͼ������<textarea rows="2" name="KSyj" cols="57"></textarea>                  
      <p>&nbsp;ǩ����<%=zg%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ʱ�䣺<%=DATe()%>  
	 <input type="hidden" name="TS" value="<%=rs("ts")%>">
    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      <p align="center">��<input type="submit" name="B1" value="�ύ����">&nbsp;                                                                           
      <input type="reset" value="ȫ������" name="B2"></p>
     </form> ��
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