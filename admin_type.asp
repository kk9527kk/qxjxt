<HTML>
<head>
<title>������</title>
<link rel="stylesheet" type="text/css" href="css.css">
<script src="bbs.js"></script>
</head>
<%if Session("username")="" Or Session("username")<>"�Ź���" then%>
<form method="POST" action="checklogin.asp" target=_self  name=form  onsubmit="return validate()">
<font color=black><center>���������Ա�ʺź����룡<br><br><center>
�û���:<input type="text" name=username size=12 class="formtext1" >&nbsp;&nbsp;
����:<input type="password" name=userpwd size=12 class="formtext1" > &nbsp;&nbsp;
<input type="hidden" name=password>
<input type="submit" name=submit value=��½ class="formtext1">&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" name=reset value=����  class="formtext1">
</font>
</form>
<%else%>
<!--#include file="conn.asp"-->


<body topmargin="0" leftmargin="0">
<center>
<table width=100% bgcolor=#ffffff border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr bgcolor=#689ACE><td><img id=shows onclick="oa_tool();" style="display:'';CURSOR: hand;" src=images/showtoc.gif id="show">
&nbsp;<img src=images/dot.gif>&nbsp;<font color=black>��ӭ���ٴλ�����<%=Session("username")%></td>

<td align=right width=70%><B><A HREF=# onclick="return load()">ˢ��</font></A> | 
 <a href=../index.asp target=_blank>����ҳ</a> | <a href=logout.asp target=_top>��ȫ�˳�</a> |</b>&nbsp;&nbsp;</td></tr>
</table>
<table width=98% bgcolor=#ffffff border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr><td><font color=red size=3><b>������</b></font></td></tr></table></center><br>


<form method="POST" action="admin_typeadd.asp" target=_self name="btypeadd">
<table width=700 border=0 align=center><tr>
<td width=70 align=center>
<input type=button value=��Ӱ�� class=formtext onclick="return benableit()">
</td>
<td width=120 align=center>
<input type="text" name="roomname" size=15 class=formtext1 disabled>
</td>
<td width=180 align=center>
<input type="text" name="blink" size=20 class=formtext1>
</td>
<td width=200 align=center>
����:
<select disabled size="1" name="broom" style="color: #689ACE" class=formtext1>
<option value="0">�ܷ�����</option>
<%
set rsads=server.createobject("adodb.recordset")
sqlads="select * from deeptree order by parentid,id"
rsads.open sqlads,conn,1,1
do while not rsads.eof
%>
<option value="<%=rsads("id")%>"><%=rsads("content")%></option>
<%
rsads.movenext
loop
%>
</select>
</td>
<td width=70 align=center>
<input disabled type=submit name=s1 value=ȷ����� class=formtext onclick="return btypeadds()">
</td>
</tr></table>
</form>

<br>
<table width=600 border=0 align=center>
<tr><td colspan="4"><font color=red>������</font></td><tr>
<td width=70 align=center><font color="#689ACE"><b>���ID</b></font></td>
<td width=130 align=center><font color="#689ACE"><b>��������</b></font></td>
<td width=130 align=center><font color="#689ACE"><b>�������</b></font></td>
<td width=130 align=center><font color="#689ACE"><b>�������</b></font></td>
<td width=100 align=center><font color="#689ACE"><b>����</b></font></td>
</tr>
<%
set rss=server.createobject("adodb.recordset")
sqls="select * from deeptree order by parentid,id"
rss.open sqls,conn,1,1
do while not rss.eof%>
<tr><form method="POST" action="admin_typeupdate.asp?id=<%=rss("id")%>" target=_self name="form<%=rss("id")%>">
<td align=center>
<input disabled type=text value="<%=rss("id")%>" size=3 name=typeid1 class=formtext1>
</td>
<td align=center>
<select size="1" name="broom1" style="color: #689ACE" class=formtext1>
<option value="0"  selected>�ܷ�����</option>
<%
set rsads=server.createobject("adodb.recordset")
sqlads="select * from deeptree where parentid=0 and parentid<>"&rss("id")&" and id<>"&rss("id")&"  order by parentid,id"
rsads.open sqlads,conn,1,1
do while not rsads.eof
%>
<option value="<%=rsads("id")%>" <%if rss("parentid")=rsads("id") then%> selected <%end if%>><%=rsads("content")%></option>
<%
rsads.movenext
loop
%>
</select>
</td><td align=center>
<input type=text value="<%=rss("content")%>" size=16 name="typename" style="color: #689ACE" class=formtext1></td>
<td align=center>
<input type=text value="<%=rss("link")%>" size=20 name="blink" style="color: #689ACE" class=formtext1></td>
<td>
&nbsp;&nbsp;<input type=submit value=�޸� class=formtext onclick="return updatetype('<%=rss("content")%>')">
&nbsp;&nbsp;<input type=button value=ɾ�� class=formtext onclick="return deltype('<%=rss("id")%>','<%=rss("content")%>')">
</td></form></tr>
<%
rss.movenext
loop
%>
</table>
</body></html>
<%end if%>