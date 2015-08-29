<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then Response.Redirect "../chat/readonly/bomb.htm"
sqlstr=Request.Form("sqlstr")
fs=int(Request.form("sqlfs"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
Set Rs=conn.Execute(sqlstr)
sqlstr=replace(sqlstr,"'","，")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','[Sql指令：]"&sqlstr&"')"
'conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','登录失败…')"
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
