<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='index.htm';}</script>"
response.end
end if

uid=request.Cookies("adminuser")
'response.write(uid)

sql="select * from sqspb  where zt='1'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰��ּ����������Ĳ���������</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="header_1.gif" leftmargin="0">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="880">
	<tr>
		<td width="878">
          <p align="center"><font color="#0000FF" face="����" size="6">
          �����ع���˰��ּ�����Ĳļ��������������</font>
        </td>
	</tr>
	<tr>
		<td width="878">
          <p align="right"><font color="#000080" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            
          </b><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b><a href="zgjmmanage.asp">:::</a></font><a href="start.asp"><font color="#000080" size="2">������ҳ</font></a><a href="zgjmmanage.asp"><font color="#000080" size="2">:::</font></a></p>
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="884"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="50" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center"><a href="index2.asp">��λ</a> 
        </td>
        <td width="60" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ������</td>
        <td width="98" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">�Ĳ�����</td>
        <td width="110" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        �ͺ�</td>
		<td width="104" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ����</td>

        <td width="116" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        <font size="2"><a href="index1.asp">��������</a></font></td>

        <%if session("admin")<>"" then%>
        <td width="62" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">��������</td>
        <%end if%>
      </tr>
      <%dim mc
    sql="select *  from SQSPB WHERE ZT='1' order by SQRQ,id "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="7" width="558"> 
          <p align="center"><font color="#FF0000">��ʱû���κ����ϼ�¼��</font></p>
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
        <td width="50" bgcolor="#FFFFFF" height="20" align="center"><a href="index1.asp?dw='<%=rs("dw")%>'"> <%=rs("dw")%></a></td>
        <td width="60" bgcolor="#FFFFFF" height="20" align="center"><a href="index1.asp?zg='<%=rs("sqr")%>'"><%=rs("sqr")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="98"><a href="index1.asp?zcmc='<%=rs("hcmc")%>'"><%=rs("hcmc")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="110"><%=rs("xh")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="104"><%=rs("sl")%></td>
        <td width="116" bgcolor="#FFFFFF" height="20" align="center">��<%=rs("sqrq")%></td>
        <%if session("admin")<>"" then%>
        <td width="62" bgcolor="#FFFFFF" height="20" align="center"><a href="edit.asp?id=<%=rs("id")%>">����</a>|<a href="del.asp?id=<%=rs("id")%>"><font color="red">ɾ��</font></a></td>
        <%end if%>
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
  <table border="0" width="886" align="center"  cellspacing="0" cellpadding="0">
    <tr>
    <td width="464">ҳ����<%=currentpage%>/<% =n%>
   	<%k=currentpage                                     
   	if k<>1 then%>
   	[<a href="index.asp?pageid=1">��ҳ</a>]                                     
   	[<a href="index.asp?pageid=<%=k-1%>">��һҳ</a>]                                     
   	<%else%>
   	[��ҳ]&nbsp;[��һҳ]                                     
   	<%end if%>
   	<%if k<>n then%>                                     
   	[<a href="index.asp?pageid=<%=k+1%>">��һҳ</a>]                                     
   	[<a href="index.asp?pageid=<%=n%>">βҳ</a>]                                     
   	<%else%>
   	[��һҳ]&nbsp;[βҳ]                                     
   	<%end if%>
        ����<font color="red"><%=totalput%></font>����¼                                    
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
          <form action="index.asp" name="pfrm">     
      <td height="23" width="418">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                 
          i=1
          do while not i > n
          %>
          <option value="index.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>�� 
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
      <td width="884" colspan="2">             
      </td>                     
    </tr>         
    <%end if%>         
    <tr>         
      <td width="884" height="35" colspan="2" valign="bottom"> 
        ��</td>        
    </tr>        
  </table>        
</div>        
  <p>
        ��
  </p>
</body>       
       
</html>       
</script>