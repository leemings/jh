<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<%
if aqjh_grade<>10 then response.redirect "../error.asp?id=440"
if aqjh_name<>application("aqjh_user") then then response.redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "delete * from letter"
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：所有的短信删除完成。');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>

