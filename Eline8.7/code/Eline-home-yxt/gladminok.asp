<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Response.Write "<body text='#000000' background='../jhimg/bk_hc3w.gif' link='#0000FF' vlink='#0000FF' alink='#0000FF'>"
grade=clng(Request.Form("grade"))
name=trim(Request.Form("name"))
if name=Application("sjjh_user") then
	Response.Write "<script Language=Javascript>alert('提示：主站长就是10级不可以更改!');location.href = 'gladmin.asp';</script>"
	response.end
end if
if sjjh_name<>Application("sjjh_user") and grade=10 then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，不可以设置10级管理员!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
select case Request.Form("submit")
case "修改"
	mp=trim(Request.Form("mp"))
	sf=trim(Request.Form("sf"))
	if sf="掌门" then
	conn.Close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：不可以设置掌门身份!');location.href = 'gladmin.asp';</script>"
	response.end
	end if
	conn.execute("update 用户 set grade="&grade&",门派='官府',身份='"&sf&"' where 姓名='"&name&"'")
case "删除"
	conn.execute("update 用户 set grade=1,门派='游侠',身份='弟子' where 姓名='"&name&"'")
case "添加"
	conn.execute("update 用户 set grade="&grade&",门派='官府',身份='弟子' where 姓名='"&name&"'")
end select
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','["&Request.Form("submit")&"]ID="&name&"["&grade&"]的管理权！')"
conn.Close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：更新完成!');location.href = 'gladmin.asp';</script>"
response.end
%>
