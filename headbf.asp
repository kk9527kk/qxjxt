<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%
uid=request.Cookies("adminuser")
sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
     zga=rs1("zg")
end if
rs1.close
%>


<head>
<style type="text/css">
<!--
.button {
	border: 1px solid #6699cc;
	font-size: 12px;
	background-color: #CCCCCC;
	height: 18px;
	width: 50px;
	margin: 1px;
	padding: 1px;
}
-->
</style>
</head>

<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#8ab4fe">
		<table width="955" border="0" cellspacing="0" cellpadding="0" bgcolor="#8ab4fe" align=center height="75">
		<tr>			
    <td width="142" height="75">¡¡</td>
	<td align=left width="807" height="75" bgcolor="#8ab4fe" background="images/logo.jpg"><font color=red>&nbsp;&nbsp;&nbsp;&nbsp;      
      </font><font color=red>&nbsp; 
      </font>
      <p><font color="red">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      </font>&nbsp;µÇÂ¼ÓÃ»§£º<%=zga%></p>
      </td> 
     

		
		</tr>
	</table>
	</body>



