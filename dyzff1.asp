<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='index.htm';}</script>"
response.end
end if
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

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="880" height="47">
	<tr>
		<td width="878" height="35">
          <p align="center"><font color="#0000FF" face="����" size="6">
          �����ع���˰��ִ�ӡֽ����</font>
          <p align="center">��
        </td>
	</tr>
	<tr>
		<td width="878" height="12">
        </td>
	</tr>
</table>
<div align="left">
        <table border="0" width="764"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000" height="77">
      <tr> 
        <td width="100" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center"> 
          <p align="center">��浥λ 
        </td>
        <td width="71" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">
        ��ӡֽ�ͺ�</td>
        <td width="94" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">�������</td>
        <td width="121" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">���ŵ�λ</td>
        <td width="216" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">
        ��������(��)</td>
		<td width="109" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">
        ���Ŵ���</td>

      
  <%dim mc
    sql="select *  from SQSPB WHERE ( qs='1' and  hcmc='��ӡֽ' and kcbz='1' and qssl>0) order by SQRQ,id "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="6" width="438" height="14"> 
          <p align="center"><font color="#FF0000">��ʱû���κδ����ŵ����ϣ�</font></p>
        </td>
      </tr>
      <%
else
const maxperpage=10 '����ÿһҳ��ʾ�����ݼ�¼�ĳ���
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
        <td width="100" bgcolor="#FFFFFF" height="36" align="center"><%=rs("dw")%></td>
        <td width="71" bgcolor="#FFFFFF" height="36" align="center"><%=rs("xh")%></td>
        <td bgcolor="#FFFFFF" height="36" align="center" width="94"><%=rs("qssl")%>��</td>
        <td bgcolor="#FFFFFF" height="36" align="center" width="121">
         <form method="POST" action="dyzffcl.asp?id=<%=rs("id")%>">
            <select size="1" name="ffdw">
              
              
              <option value="�ֳ���">�ֳ���</option>
            <option value="�칫��">�칫��</option>
            <option value="�����">�����</option>
            <option value="˰����">˰����</option>
              <option value="����˰��">����˰��</option>
            <option value="���ܿ�">���ܿ�</option>
            <option value="�Ʋƿ�">�Ʋƿ�</option>
            <option value="�˽̿�">�˽̿�</option>
            <option value="�����">�����</option>
            <option value="��Ϣ����">��Ϣ����</option>
            <option value="��˰��">��˰��</option>
            <option value="�����">�����</option>
            <option value="����һ��">����һ��</option>
            <option value="��������">��������</option>
            <option value="������">������</option>
            <option value="ƽ����">ƽ����</option>
            <option value="��԰��">��԰��</option>
            <option value="������">������</option>
            <option value="��¡��">��¡��</option>
            <option value="������">������</option>
              
              
              </select>
            ��
           

</td>
        <td bgcolor="#FFFFFF" height="36" align="center" width="216"><input type="text" name="ffSL" size="12">
        </td>
        <td bgcolor="#FFFFFF" height="36" align="center" width="109"><input type="submit" value="����" name="B1"></td>   </form>
       
        </form> 
      </tr>
       <%
      i=i+1
    rs.movenext
    loop
    end if
    %>

    </table>
<table border="0" width="764" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td width="464">ҳ����<%=currentpage%>
      /<% =n%>                                                                                             
      <%k=currentpage                                                                                                                                                            
   	if k<>1 then%>
      [<a href="sq.asp?pageid=1">��ҳ</a>] [<a href="sq.asp?pageid=<%=k-1%>">��һҳ</a>]                                                                                              
      <%else%>
      [��ҳ]&nbsp;[��һҳ] <%end if%>                                                                                             
      <%if k<>n then%>                                                                                             
      [<a href="sq.asp?pageid=<%=k+1%>">��һҳ</a>] [<a href="sq.asp?pageid=<%=n%>">βҳ</a>]                                                                                              
      <%else%>
      [��һҳ]&nbsp;[βҳ] <%end if%>                                                                                             
      ����<font color="red"><%=totalput%>                                                                                             
      </font>����¼ <b><font size="2" color="#0000FF">&nbsp;</font></b></td>                                                                                             
    <td height="23" width="302">
    <form action="sq.asp" name="pfrm">   
      <p align="right"><select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
        <%                                                                                                                                        
          i=1
          do while not i > n
          %>
        <option value="sq.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>�� 
        <%=i%> ҳ</option>
        <%
          i=i+1
          loop
          %>
      </select></td>                                                                                             
  </tr></form>
</table>

</html>