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
<title>�����ع���˰��ּ�����Ĳļ����������������</title>
</head>

<body background="images/header_1.gif">
<%sql="select * from QXJDJb  where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if  rs.eof then
    response.write "<script language=JavaScript>{window.alert('�Բ���,û�ҵ�����Ϣ!');window.location.href='XH.asp';}</script>"
    response.end
else

%>
<p align="center"><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>       
<table align="center" border="1" width="750" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#FFFFFF" height="232">
  <tr>
    <td width="811" bordercolorlight="#000000" height="36" background="file:///D:/qxjxt/images/bg1.gif" colspan="2">
      <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰���ְ��������������</font></td>
  </tr>
  <tr>
    <td width="160" bordercolorlight="#000000" height="36">
      <p align="center">�����ˣ�<%=rs("zg")%>
    </td>
    <td width="751" bordercolorlight="#000000" height="36">���ʱ���� �� <font color="#FF00FF">&nbsp;<%=rs("qjsjq")%>           
      </font>&nbsp;&nbsp;&nbsp; ���ʱ��ֹ�� <font color="#FF00FF">&nbsp; <%=rs("qjsjz")%>           
      </font> &nbsp;�����<font color="#FF00FF">&nbsp;<%=rs("ts")%>
      </font><font color="#000000">��</font></td>       
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">������ɣ�(<%=rs("qjsx")%>
      )</td>       
  </tr>
  <tr>
    <td width="268" bordercolorlight="#000000" height="36">������(<font color="#FF00FF">&nbsp;<%=rs("qjLB")%>
      </font>)</td>
    <% 
   SFLFJ=RS("SFLFJ")
   IF SFLFJ="YES" THEN
     SFLFJ1="��"
   ELSE 
     SFLFJ1="��"
   END IF  
%>
<td width="539" bordercolorlight="#000000" height="36">ǰ���εأ�(<font color="#FF00FF">&nbsp;<%=rs("qWHD")%>
      </font>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��ֳ������־��(<font color="#FF00FF">&nbsp;<%=SFLFJ1%>       
      </font>)</td>
  </tr>
  <tr>
    <td width="807" bordercolorlight="#000000" height="36" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ڣ�<font color="#FF00FF">&nbsp;<%=RS("SQSJ")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>       
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">��λ�쵼���������<font color="#FF00FF">&nbsp;(<%=rs("ksyj")%>
      )</font>       
      <p>&nbsp;ǩ����<font color="#FF00FF">&nbsp;<%=RS("KSQM")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        
      ʱ�䣺<font color="#FF00FF">&nbsp;<%=RS("DWRQ")%>
      </font></td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">��١�̽�ס��飨ɥ�������߹涨���˽̿������        
      (<font color="#FF00FF">&nbsp;<%=RS("RJYJ")%>
      </font>)
      <p>&nbsp;ǩ����<font color="#FF00FF">&nbsp;<%=RS("RJ")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      ǩ��ʱ�䣺<font color="#FF00FF">&nbsp;<%=rs("rjkrq")%>
      </font></td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">�ֹܾ��쵼���������        
      (<font color="#FF00FF">&nbsp;<%=rs("fgldyj")%>
      </font>)<p>&nbsp;ǩ����<font color="#FF00FF">&nbsp;<%=rs("fgld")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      ǩ��ʱ�䣺<font color="#FF00FF">&nbsp;<%=rs("fgldrq")%>
      </font></td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">�ؾ־ֳ������(<font color="#FF00FF">&nbsp;<%=rs("JZYJ")%></font>) 
      <p>&nbsp;ǩ����<font color="#FF00FF">&nbsp;<%=rs("JZ")%>
      </font>      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      ǩ��ʱ�䣺<font color="#FF00FF">&nbsp;<%=rs("JZrq")%>
      </font>    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">�����������ʵ������(<font color="#FF00FF">&nbsp;<%=rs("sxts")%>
      </font>)<p>&nbsp;�����ˣ�<font color="#FF00FF">&nbsp;<%=rs("XJR")%>
      </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      �������ڣ�<font color="#FF00FF">&nbsp;<%=rs("XJRQ")%>
      </font>&nbsp;    </td>
  </tr>
  <tr>
    <td width="811" bordercolorlight="#000000" height="1" colspan="2">
      <p align="center">��&nbsp; <a href="javascript:history.go(-1);">����</a>  
     
      </p>
      </form>��</td>
  </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="66%" height="55">
  <tr>
    
  </tr>
</table>
</body>
<%end if %>
</html>





