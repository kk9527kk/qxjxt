<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='index.htm';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
'response.write(uid)

sql="select * from ZG  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("zw")
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰��������ٹ���ϵͳ</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="769">
	<tr>
		<td width="767">
          <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰�����������������</font>
          <p align="center">��
        </td>
	</tr>
	<tr>
		<td width="767">
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="772"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="50" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center">��λ 
        </td>
        <td width="60" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        �����</td>
        <td width="60" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ���ʱ����</td>
        <td width="68" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">���ʱ��ֹ</td>
        <td width="49" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        �������</td>
		<td width="54" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ����ʱ��</td>

        <td width="147" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        <font size="2">���״̬</a></font></td>

       
        <td width="60" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">��������</td>
       
      </tr>
      <%dim mc
      
      
    'sql="select *  from QXJDJB WHERE QJZT='3' AND   '"&DW1&"' LIKE  '%"&DW&"%'   order by SQSJ"
    sql="select * from qxjdjb inner join dw on qxjdjb.dw=dw.dw  where qxjdjb.qjzt='3' and dw.fgld='"&zg&"'"
       set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="8" width="446"> 
          <p align="center"><font color="#FF0000">��ʱû���κδ����������ϣ�</font></p>
        </td>
      </tr>
      <%
else
const maxperpage=15 '����ÿһҳ��ʾ�����ݼ�¼�ĳ���
dim currentpage '���嵱ǰҳ�ı���
rs.pagesize=maxperpage
currentpage=request.querystring("pageid")
if currentpage="" then
currentpage=1
elseif currentpage<1 then
currentpage=1
else
currentpage=clng(currentpage)
	if currentpage > rs.pagecount then
	currentpage=rs.pagecount
	end if
end if
'�������currentpage���������Ͳ�����ֵ��
'��1��������currentpage
if not isnumeric(currentpage) then
currentpage=1
end if
dim totalput,n '�������
totalput=rs.recordcount
if totalput mod maxperpage=0 then
n=totalput\maxperpage
else
n=totalput\maxperpage+1
end if
if n=0 then
n=1
end if
rs.move(currentpage-1)*maxperpage
i=0
do while i< maxperpage and not rs.eof
%>
      <tr> 
        <td width="50" bgcolor="#FFFFFF" height="20" align="center"> <%=rs("qxjdjb.dw")%></td>
        <td width="60" bgcolor="#FFFFFF" height="20" align="center"><%=rs("ZG")%></td>
        <td width="60" bgcolor="#FFFFFF" height="20" align="center"><%=rs("QJSJQ")%>��</td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="68"><%=rs("QJSJZ")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="49"><%=rs("TS")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="54"><%=rs("SQSJ")%></td>
        <%
        QJZT=RS("QJZT")
        SELECT CASE QJZT
           CASE "1"
             QJZT1="�����������������"
           CASE "2"
             QJZT1="������������˽̿Ƹ���"
           CASE "3"
             QJZT1="������������ֳֹܾ�����"
           CASE "4"
             QJZT1="�ֳֹܾ�������ֳ�����"
           CASE "5"
             QJZT1="�ֳ�����ͬ��"
            CASE "6"
             QJZT1="�˽̿�ͬ������"
           CASE "7"
             QJZT1="�ֳֹܾ�ͬ������"
           CASE "8"
             QJZT1="�˽̿�ͬ����ֳ�����"
           CASE "9"
             QJZT1="��ͬ��"
			CASE "10"
			 QJZT1="������ͬ�Ⲣ����"
         END SELECT            
        %>
        
        <td width="147" bgcolor="#FFFFFF" height="20" align="center">��<%=QJZT1%></td>
      
        <td width="60" bgcolor="#FFFFFF" height="20" align="center"><a href="SPJLDk.asp?id=<%=rs("qxjdjb.id")%>"><font color="#FF0000">����</font></a></td>
       
        <%
      i=i+1
    rs.movenext
    loop
    end if
    %>
      </tr>
    </table>
</div>
<div align="left">
  <table border="0" width="771" align="center"  cellspacing="0" cellpadding="0">
    <tr>
    <td width="464">ҳ����<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                    
   	if k<>1 then%>
   	[<a href="SPJLD.asp?pageid=1">��ҳ</a>]                                                    
   	[<a href="SPJLD.asp?pageid=<%=k-1%>">��һҳ</a>]                                                    
   	<%else%>
   	[��ҳ]&nbsp;[��һҳ]                                                    
   	<%end if%>
   	<%if k<>n then%>                                                    
   	[<a href="SPJLD.asp?pageid=<%=k+1%>">��һҳ</a>]                                                    
   	[<a href="SPJLD.asp?pageid=<%=n%>">βҳ</a>]                                                    
   	<%else%>
   	[��һҳ]&nbsp;[βҳ]                                                    
   	<%end if%>
        ����<font color="red"><%=totalput%></font>����¼                                                   
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
          <form action="SPJLD.asp" name="pfrm">     
      <td height="23" width="303">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                              
          i=1
          do while not i > n
          %>
          <option value="SPJLD.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>�� 
          <%=i%> ҳ</option>
          <%
          i=i+1
          loop
          %>
        </select>                              

        </td>
    </tr>
    </form>
  <center>
    <%if session("admin")<>"" then%>
    <tr>
      <td width="769" colspan="2">             
      </td>                     
    </tr>         
    <%end if%>         
    <tr>         
      <td width="769" height="35" colspan="2" valign="bottom"> 
        ��</td>        
    </tr>        
  </table>        
</div>        
  <p>
        ��
  </p>
</center>
</body>       
       
</html>       
</script>




































