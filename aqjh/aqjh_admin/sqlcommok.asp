<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
sqlstr=Request.Form("sqlstr")
fs=int(Request.form("sqlfs"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Set Rs=conn.Execute(sqlstr)
sqlstr=replace(sqlstr,"'","，")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','[Sql指令：]"&sqlstr&"')"
'conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','登录失败…')"
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：所执行的指定执行完成！点确定反回！');</script>"
%>
<script language="vbscript">
location.href = "javascript:history.go(-1)"
</script>
<%
response.end
%>
