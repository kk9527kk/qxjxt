<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<%
response.contentType="text/xml"
response.expires=0
dim conn,rs,xmlstr,nodeId
nodeId=trim(request.querystring("id"))

Function coder(str)
 Dim result,L,i
 If IsNull(str) Then : coder="" : Exit Function : End If
 L=Len(str) : result=""
For i = 1 to L
  select case mid(str,i,1)
	case "<"     : result=result+"&lt;"
	case ">"     : result=result+"&gt;"
	case chr(34) : result=result+"&quot;"
	case "&"     : result=result+"&amp;"
	case chr(13) : result=result+"<br/>"
	case chr(9)  : result=result+"&nbsp;&nbsp;"
	case chr(32) : result=result+"&nbsp;"
	case else    : result=result+mid(str,i,1)
  end select
Next
 coder=result
End Function 


xmlstr="<?xml version='1.0' encoding='gb2312'?>"&vbCrLf
xmlstr=xmlstr&"<xml>"&vbCrLf

if nodeId<>"" and isNumeric(nodeId) then
''论坛版块
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "select *,(select count(*) from deeptree where parentid = T.id and level>="&CInt(session("level"))&") as children from deeptree T where parentid="&CInt(nodeId)&" and not(link<>null) and level>="&CInt(session("level"))&" order by id",conn,1,3

if not rs.eof then
do while not rs.eof
	xmlstr=xmlstr&"<TreeNode id='"&rs("id")&"'>"&vbCrLf
	xmlstr=xmlstr&"<NodeText>"&coder(trim(rs("content")))&"</NodeText>"&vbCrLf
	xmlstr=xmlstr&"<title></title>"&vbCrLf
	xmlstr=xmlstr&"<NodeUrl>"&coder("index1.asp?type=id&id="&rs("id"))&"</NodeUrl>"&vbCrLf
	xmlstr=xmlstr&"<child>"&rs("children")&"</child>"&vbCrLf
	xmlstr=xmlstr&"<target></target>"&vbCrLf
	xmlstr=xmlstr&"</TreeNode>"&vbCrLf
rs.movenext
loop
end if
''系统设置版块
set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open "select *,(select count(*) from deeptree where parentid = T.id and level>="&CInt(session("level"))&") as children from deeptree T where parentid="&CInt(nodeId)&" and link<>'' and level>="&CInt(session("level")),conn,1,3

if not rs1.eof then
do while not rs1.eof
	xmlstr=xmlstr&"<TreeNode id='"&rs1("id")&"'>"&vbCrLf
	xmlstr=xmlstr&"<NodeText>"&coder(trim(rs1("content")))&"</NodeText>"&vbCrLf
	xmlstr=xmlstr&"<title></title>"&vbCrLf
	if rs1("link")="nolink" then
	xmlstr=xmlstr&"<NodeUrl></NodeUrl>"&vbCrLf
	else
	xmlstr=xmlstr&"<NodeUrl>"&coder("index1.asp?type=url&url="&rs1("link"))&"</NodeUrl>"&vbCrLf
	end if
	xmlstr=xmlstr&"<child>"&rs1("children")&"</child>"&vbCrLf
	xmlstr=xmlstr&"<target></target>"&vbCrLf
	xmlstr=xmlstr&"</TreeNode>"&vbCrLf
rs1.movenext
loop
end if

set rs=nothing
set rs1=nothing
set conn=nothing
end if

xmlstr=xmlstr&"</xml>"
response.write xmlstr
%>