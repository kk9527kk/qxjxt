<!--#include file="conn.asp"-->
<html>
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
lid=request.QueryString("id")
'response.write(lid)
if  request.Cookies("adminuser")="" then
   response.write "��Դδ֪�����ݶ�ʧ!"
   response.end
end if
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰����˺Ź���</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1" align="center" cellpadding="0" width="713">
	<tr>
		<td width="711">
          <p align="center"><font color="#0000FF" face="����" size="6">
          �����ع���˰����˺Ź���</font><p align="center">��
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="715" align="center" cellpadding="0" bgcolor="#000000" cellspacing="1">
      <tr> 
        <td width="142" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center">�˺� 
        </td>
        <td width="142" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        <p align="center">
        ��λ</p>
        </td>
        <td width="143" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">����</td>
        <td width="143" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ���</td>
        <td width="143" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">����</td>
              </tr>
      <%dim mc
    sql="select *  from zg order by  id  "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="5" width="352"> 
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
        <td width="142" bgcolor="#FFFFFF" height="20" align="center"><%=rs("id")%></td>
        <td width="142" bgcolor="#FFFFFF" height="20" align="center"><%=rs("dw")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="143"><%=rs("zg")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="143"><%=rs("xh")%></td>
        <td width="143" bgcolor="#FFFFFF" height="20" align="center"><a href="ryedit.asp?id=<%=rs("id")%>">�༭</a>|<a href="rydel.asp?id=<%=rs("id")%>"><font color="red">ɾ��</font></a></td>
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
  <table border="0" width="713" align="center" cellpadding="0">
    <tr>
    <td width="464" align="left">ҳ����<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                
   	if k<>1 then%>
   	[<a href="zhgl.asp?pageid=1">��ҳ</a>]                                                
   	[<a href="zhgl.asp?pageid=<%=k-1%>">��һҳ</a>]                                                
   	<%else%>
   	[��ҳ]&nbsp;[��һҳ]                                                
   	<%end if%>
   	<%if k<>n then%>                                                
   	[<a href="zhgl.asp?pageid=<%=k+1%>">��һҳ</a>]                                                
   	[<a href="zhgl.asp?pageid=<%=n%>">βҳ</a>]                                                
   	<%else%>
   	[��һҳ]&nbsp;[βҳ]                                                
   	<%end if%>
        ����<font color="red"><%=totalput%></font>����¼                                               
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
          <form action="zhgl.asp" name="pfrm">     
      <td height="23" width="245" align="left">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                            
          i=1
          do while not i > n
          %>
          <option value="zhgl.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>�� 
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
       <tr>
      <td width="711" colspan="2" align="left"><a href="ryadd.asp"><font color="#000080">����˺���Ϣ</font></a><font color="#000080">   
 </font>                        
      </td>                     
    </tr>         
          
    <tr>         
      <td width="711" height="35" colspan="2" valign="bottom" align="left"> 
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