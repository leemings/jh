<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<%
if aqjh_grade<>10 then response.redirect "../error.asp?id=440"
id=request("id")
if not isnumeric(id) then
	response.write "id��ʹ������"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "delete * from letter where id="&id
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��idΪ"&id&"�Ķ���ɾ����ɡ�');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>

