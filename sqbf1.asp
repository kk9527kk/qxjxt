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
sql="select * from djb  where id='"&uid&"'"
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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰��ּ�����Ĳļ��������������</title>
</head>

<body background="images/header_1.gif" bgproperties="fixed">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="880">
	<tr>
		<td width="878">
          <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰��ּ�����Ĳļ����ǩ�մ���</font>
        </td>
	</tr>
	<tr>
		<td width="878">
          <p align="right"><b><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><font color="#000080"> 
          </font>  
          </b><font color="#000080">��ǰ�û�:<%=zg%><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b><a href="modifypass.asp">:::�޸�����:::</a></font><b>&nbsp;</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#000080" size="2"> 
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><a href="login.asp"><font color="#000080">:::���µ�¼:::</font></a></p>          
        </td>
	</tr>
</table>
    <table border="0" width="884"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="83" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center">�������� 
        </td>
        <td width="87" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ������</td>
        <td width="95" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">ǩ����</td>
        <td width="118" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">�Ĳ�����</td>
        <td width="231" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        �ͺ�</td>
		<td width="119" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        ��������</td>

        <td width="171" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        �Ƿ�����</td>

       
        <td width="150" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">ǩ������</td>
       
       
        <td width="146" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">ǩ�ս��</td>
       
       
        <td width="84" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">����</td>
       
      
      
  <%dim mc
    sql="select *  from SQSPB WHERE dw='"&dw&"' order by id desc "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="10" width="547"> 
          <p align="center"><font color="#FF0000">��ʱû���κδ����������ϣ�</font></p>
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
        <td width="83" bgcolor="#FFFFFF" height="20" align="center"><%=rs("sqrq")%></a></td>
        <td width="87" bgcolor="#FFFFFF" height="20" align="center"><%=rs("sqr")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="95"><%=rs("qsr")%>��</td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="118"><%=rs("hcmc")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="231"><%=rs("xh")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="119"><%=rs("sl")%></td>
        <td width="171" bgcolor="#FFFFFF" height="20" align="center">
        
        <%if rs("zt")="4" then %>
        <font color="#FF0000">����ͬ��</font>  
        <%elseif rs("zt")="5" or rs("zt")="6" or rs("zt")="7"  then%>  
        <font color="#008000">����ͬ�� </font>   
        <%else%>
         <font color="#008880">δ���� </font>   
        <%end if%>
        </td>
        <td width="150" bgcolor="#FFFFFF" height="20" align="center">
         <%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
               <form method="POST" action="qs.asp?id=<%=rs("id")%>" NAME="FORM1">
            <input type="text" name="QSSL" size="15" value="<%=RS("SL")%>">
          <%else%>  
            <%=RS("SL")%>
         <%END IF%>                                ��</td>
  
        <td width="146" bgcolor="#FFFFFF" height="20" align="center">
             <p align="center">
             <%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
        <input type="text" name="QSJE" size="13" value="<%=RS("JE")%>"></p>
         </td>
       <%END IF%>
        <td width="84" bgcolor="#FFFFFF" height="20" align="center" valign="middle">
        
        <%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
        <input type="submit" value="ǩ��" name="B1">   
         <%elseif rs("zt")="4" and rs("qs")="1"  then%>                             
         <font color="#008880">���� </font>                                
         <%else%>
         <font color="#008880"></font>                                
        <%end if%>
              
        
        </td>
       </form> 
      </tr>
       <%
      i=i+1
    rs.movenext
    loop
    end if
    %>

    </table>
<table border="0" width="880" align="center" cellspacing="0" cellpadding="0">
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
    <td height="23" width="418">
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
<p align="center">                          
��</p>
<p align="center">                          
<form method="POST" action="save.asp?id=<%=uid%>" name="form1">
<table align="center" border="1" width="880" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#FFFFFF" height="232">
  <tr>
    <td width="876" bordercolorlight="#000000" height="36" background="images/bg1.gif">
      <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰��ּ�����Ĳļ��������</font>
     
    </td>
  </tr>
  <tr>
    <td width="876" bordercolorlight="#000000" height="36">
      <p align="center">                          
���뵥λ:<%=dw%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                                                             
������:<%=zg%>
     
    </td>
  </tr>
  <tr>
    <td width="872" bordercolorlight="#000000" height="36">����Ĳ���:<%if dw="��Ϣ����"  then%><select size="1" name="hcmc">
        <option value="����">����</option>
        <option value="ī��">ī��</option>
        <option value="ɫ��">ɫ��</option>
        <option value="ɫ����">ɫ����</option>
        <option value="����">����</option>
        <option value="�ƶ�Ӳ��">�ƶ�Ӳ��</option>
        <option value="�������">�������</option>
        <option value="��ӡֽ">��ӡֽ</option>
        <option value="����">����</option>
        <%else%>
      </select><select size="1" name="hcmc">
         <option value="����">����</option>
        <option value="ī��">ī��</option>
        <option value="ɫ��">ɫ��</option>
        <option value="ɫ����">ɫ����</option>
        <option value="����">����</option>
        <option value="�ƶ�Ӳ��">�ƶ�Ӳ��</option>
        <option value="�������">�������</option>
        <option value="����">����</option>

        <%end if %>
        &nbsp;                                                                              
        </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                     
      �Ĳ��ͺ�:<input type="text" name="xh" size="24">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                     
      ����:<input type="text" name="SL" size="9" value="1">
     
    </td>
  </tr>
  <tr>
    <td width="876" bordercolorlight="#000000" height="45">��������:<textarea rows="2" name="sqsx" cols="114"></textarea></td>
  </tr>
  <tr>
    <td width="876" bordercolorlight="#000000" height="1">
      <p align="center">��<input type="submit" name="B1" value="�ύ����">&nbsp;                                                        
      <input type="reset" value="ȫ����д" name="B2"></p>
      ��
    </td>
  </tr>
</table>
 </form>
</body>

</html>


