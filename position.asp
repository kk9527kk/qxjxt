<%@ Language=VBScript%>
<%
response.contentType="text/xml"
response.charSet="GB2312"
response.expires=0
dim id,xmlstr,conn
id=trim(request.querystring("id"))

Function getPosition(uid)
	set rs=conn.execute("select * from deeptree where id="&uid&" and level>='"&CInt(session("level"))&"' order by id")
	if not rs.eof then
	rs.movefirst
	xmlstr=xmlstr&"<node>"&rs("parentid")&"</node>"&vbCrLf
	getPosition(rs("parentid"))
	end if
	set rs=nothing
end Function

xmlstr="<?xml version='1.0' encoding='gb2312'?>"&vbCrLf
xmlstr=xmlstr&"<xml>"&vbCrLf

if id<>"" and isNumeric(id) then
%>
<!--#include file="conn.asp"-->
<%
getPosition(id)
set conn=nothing
end if

xmlstr=xmlstr&"</xml>"
response.write xmlstr

%>