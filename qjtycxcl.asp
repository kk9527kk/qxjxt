<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
'lid=request.QueryString("id")
    uid=request.Cookies("adminuser")
    response.cookies("flag") = "loginok"
    response.cookies("adminuser") = uid
    rqq=request.Cookies("rqq")
    rqz=request.Cookies("rqz")
    dw1=request.Cookies("dw")
    zg1=request.Cookies("zg")

%>

<html>
<%
 response.cookies("rqq") = rqq
 response.cookies("rqz") = rqz

rqq1=FormatDateTime(rqq,1)
rqz1=FormatDateTime(rqz,1)

%>


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

<body topmargin="2" background="IMAGES/header_1.gif" leftmargin="0">
<table border="0" id="table1" align="center"  cellspacing="0" cellpadding="0" width="759">
	<tr>
		<td width="757">
          <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰���ְ�������ٲ�ѯ�嵥</font>
        </td>
	</tr>
	<tr>
		<td width="757">
<p align="center"><b>����ʱ��:��<%=rqq1%>��<%=rqz1%>        </b>        </p>
        </td>
	</tr>
	<tr>
		<td width="757">
<br>
<p>��
        </td>
	</tr>
</table>

    <table border="0" width="775"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
         <td width="49" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        <font color="#008080">
        ���</font></td>
        <td width="77" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">������</td>
        <td width="70" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">�ݼ�ʱ����</td>
        <td width="106" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">�ݼ�ʱ��ֹ</td>
        <td width="90" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">�����������</td>

		<td width="194" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center"> ����״̬</td>

		<td width="88" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center"> ��������</td>

             </tr>
            <% 
            
        if  zg1="" then
          bds="dw='"&dw1&"'"& "and sqsj between #"&rqq&"#"&" and #"&rqz&"#"  
        else
          bds="dw='"&dw1&"'"& " and zg='"&zg1& "' and sqsj between #"&rqq&"#"&" and #"&rqz&"#"  
        end if    
         
          sql="select *  from qxjdjb  where "&bds &" order by zg desc"
          set rs=server.createobject("adodb.recordset")
          rs.open sql,connstr,1,1
       if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="7" width="769"> 
          <p align="center"><font color="#FF0000">��ʱû���κ����ϣ�</font></p>
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
         <td width="49" bgcolor="#FFFFFF" height="23" align="center"> 
         
          <% 
          qjzt=rs("qjzt")
          
        SELECT CASE QJZT
           CASE "1"
             QJZT1="�����������������"
           CASE "2"
             QJZT1="������������˽̿Ƹ���"
           CASE "3"
             QJZT1="�˽̿�������ֳֹܾ�����"
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
         END SELECT 
                     %>��
 
         
         
         
         
       <a href="cx.asp?id=<%=rs("id")%>"> <font color="#008080">
           <%=rs("id")%></font></a></td>
        <td width="77" bgcolor="#FFFFFF" height="23" align="center">    <%=rs("zg")%>��</td>
        <td width="70" bgcolor="#FFFFFF" height="23" align="center">��  <%=rs("qjsjq")%></td>
        <td width="106" bgcolor="#FFFFFF" height="23" align="center">   <%=rs("qjsjz")%>��</td>
        <td width="90" bgcolor="#FFFFFF" height="23" align="center">��  <%=rs("sqsj")%></td>
		<td width="194" bgcolor="#FFFFFF" height="23" align="center">��<%=QJZT1%> </td>
		<td width="88" bgcolor="#FFFFFF" height="23" align="center">��  <%=rs("XJRQ")%> </td>

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
      [<a href="qjtycxcl.asp?pageid=1">��ҳ</a>] [<a href="qjtycxcl.asp?pageid=<%=k-1%>">��һҳ</a>]                                                                                                                   
      <%else%>
      [��ҳ]&nbsp;[��һҳ] <%end if%>                                                                                                                  
      <%if k<>n then%>                                                                                                                  
      [<a href="qjtycxcl.asp?pageid=<%=k+1%>">��һҳ</a>] [<a href="qjtycxcl.asp?pageid=<%=n%>">βҳ</a>]                                                                                                                   
      <%else%>
      [��һҳ]&nbsp;[βҳ] <%end if%>                                                                                                                  
      ����<font color="red"><%=totalput%>                                                                                                                  
      </font>����¼ <b><font size="2" color="#0000FF">&nbsp;</font></b></td>                                                                                                                  
    <td height="23" width="302">
     
      <p align="right"><select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
        <%                                                                                                                                                           
          i=1
          do while not i > n
          %>
        <option value="qjtycxcl.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>�� 
        <%=i%> ҳ</option>
        <%
          i=i+1
          loop
        rs.close
        set rs=nothing

  %>
      </select></td>                                                                                                                
  </tr>
</table>


<center>
</center>

</body>       
       
</script>




















































































