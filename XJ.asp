

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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰��������ٹ���</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

&nbsp;           

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->            

<%            
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
sql="select * from ZG  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("ZW")
end if
rs1.close
%>

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="753">
	<tr>
		<td width="751">
          <p align="center"><font color="#0000FF" face="����" size="6">
          �����ع���˰������ٴ���</font>
        </td>
	</tr>
	<tr>
		<td width="751">
          ��
          <p>��
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="761"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="85" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center"><a href="index2.asp">��λ</a> 
        </td>
        <td width="93" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ������</td>
        <td width="99" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">���ʱ����</td>
        <td width="150" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ���ʱ��ֹ</td>
		<td width="120" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        �������</td>

        <td width="118" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ����ʱ��</td>

       
        <td width="78" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">ʵ������</td>
       
       
        <td width="95" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">����</td>
       
      </tr>
      <%dim mc
    	select case ZW
		case "3"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and dw='"&dw&"' order by SQSJ "
		case "6"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and zw=3 order by SQSJ "
		case "7"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and zw=3 order by SQSJ "
		case "5"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and zw=4 order by SQSJ "
		end select
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="8" width="719"> 
          <p align="center"><font color="#FF0000">��ʱû���κδ����ٵ����ϣ�</font></p>
        </td>
      </tr>
      <%
else
const maxperpage=25 '����ÿһҳ��ʾ�����ݼ�¼�ĳ���
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
        <td width="85" bgcolor="#FFFFFF" height="20" align="center"><%=rs("dw")%></a></td>
        <td width="93" bgcolor="#FFFFFF" height="20" align="center"><%=rs("ZG")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="99"><%=rs("QJSJQ")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="150"><%=rs("QJSJZ")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="120"><%=rs("TS")%></td>
        <td width="118" bgcolor="#FFFFFF" height="20" align="center"><a href="cx.asp?id=<%=rs("id")%>"><font color="#FF0000">&nbsp;&nbsp;<%=rs("SQSJ")%></font></a></td>
      
        <td width="78" bgcolor="#FFFFFF" height="20" align="center">
       <form method="POST" action="XJCL.asp?id=<%=rs("id")%>" name="form1">
           <input type="text" name="SXTS" size="7" value="<%=rs("TS")%>">��
          </td>
       
        <td width="95" bgcolor="#FFFFFF" height="20" align="center" valign="middle"><font color="#FF0000"><input type="submit" value="����" name="B1"></font></form></td>
       
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
  <table border="0" width="762" align="center"  cellspacing="0" cellpadding="0">
    <tr>
    <td width="464">ҳ����<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                                                
   	if k<>1 then%>
   	[<a href="XJ.asp?pageid=1">��ҳ</a>]                                                                                
   	[<a href="XJ.asp?pageid=<%=k-1%>">��һҳ</a>]                                                                                
   	<%else%>
   	[��ҳ]&nbsp;[��һҳ]                                                                                
   	<%end if%>
   	<%if k<>n then%>                                                                                
   	[<a href="XJ.asp?pageid=<%=k+1%>">��һҳ</a>]                                                                                
   	[<a href="XJ.asp?pageid=<%=n%>">βҳ</a>]                                                                                
   	<%else%>
   	[��һҳ]&nbsp;[βҳ]                                                                                
   	<%end if%>
        ����<font color="red"><%=totalput%></font>����¼                                                                               
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
            
      <td height="23" width="294">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                                                            
          i=1
          do while not i > n
          %>
          <option value="XJ.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>�� 
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
      <td width="760" colspan="2">             
      </td>                     
    </tr>         
    <%end if%>         
    <tr>         
      <td width="760" height="35" colspan="2" valign="bottom"> 
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