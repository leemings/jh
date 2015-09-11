<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
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
username=Request.Form("search")
cz=Request.form("sjcz")
name=request.querystring("username")
if name<>"" then
username=name
cz="姓名"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if cz="ID" then
username=int(username)
sqlstr="SELECT * FROM 用户 where "&cz&"="&username
rs.Open sqlstr,conn
else
sqlstr="SELECT * FROM 用户 where "&cz&"='"&username&"'"
rs.Open sqlstr,conn
end if
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要查找的人我们找不到！请查看是否正确！');history.back();</script>"
else
if rs("grade")=10 and aqjh_name<>Application("aqjh_user") then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：不可以修改站长资料！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','查找：["&cz&"]="&username&"的记录！')"
Response.Write "<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666><LINK href=css/css.css type=text/css rel=stylesheet>"
Response.Write "<form method=post action=updateuser.asp><table  border=0 cellspacing=1 cellpadding=4 bgcolor=#f2f2ea height='20' align='center'><tr><td colspan='6'><div align='center'>江湖ID:"&rs("id")&"<input type=hidden readonly size=4 name=id value="&rs("id")&">(注意：数据类型与所对应的数据必需相同，否则出错！不能有空的数据存在!)</div></td></tr>"
Response.Write "<tr>"
for i=1 to rs.Fields.Count-25
strtype=rs.fields(i).Type
strname=rs.Fields(i).Name
if strname="姓名" then seename=rs.Fields(i).value
if rs.fields(i).Type=202 then strtype="字符"
if rs.fields(i).Type=3 then strtype="数字"
if rs.fields(i).Type=7 then strtype="日期"
if rs.fields(i).Type=11 then strtype="逻辑"
if strname="grade" then strname="管理等级"
if strname="mvalue" then strname="月积分"
if strname="allvalue" then strname="总积分"
if strname="times" then strname="登录次数"
if strname="regip" then strname="注册ip"
if strname="lastip" then strname="最后ip"
if strname="regtime" then strname="注册时间"
if strname="lasttime" then strname="最后时间"
'if strname<>"cw" and strname<>"w1" and strname<>"w2"  and strname<>"w3"  and strname<>"w4"  and strname<>"w5"  and strname<>"w6"    and strname<>"w7"   and strname<>"w8" and strname<>"z1"   and strname<>"z2"   and strname<>"z3"   and strname<>"z4"   and strname<>"z5"   and strname<>"z6"  then
	Response.Write "<td bgcolor='#f2f2ea'><div align='right'>"&strname&"(<font color=blue>"&strtype&"</font>)：<input type=text size='10' name="&rs.Fields(i).Name&" value='"&rs.Fields(i).value&"' class=form style='BORDER-BOTTOM:#B8AF86 1px solid'></font></div></td>"
	if i/3=int(i/3) then Response.Write "</tr><tr>"
'end if
	next
Response.Write "</tr>"
Response.Write "<tr><td colspan='6'><div align='center'><input type=submit value=更新 name=submit> <input type=submit value=新增 name=submit> <input type=submit value=删除 name=submit> <input type=reset value=重置> <a href='../chat/TOWUPIN.asp?toname="&seename&"' target='_blank' >查询此人物品</a></div></td></tr>"
Response.Write "</table></form>"
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
