<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="update 用户 set grade=1 where grade<=5 and 身份<>'掌门'"
Set Rs=conn.Execute(sql)
sql="update 用户 set grade=5 where grade<=6 and 身份='掌门'"
Set Rs=conn.Execute(sql)
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','普通用户管理为1级，掌门管理为5级(管理员不变)操作完成!')"
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('普通用户管理为1级，掌门管理为5级(管理员不变)操作完成!');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>
