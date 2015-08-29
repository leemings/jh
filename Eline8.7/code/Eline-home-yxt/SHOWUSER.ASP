<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
username=Request.Form("search")
cz=Request.form("sjcz")
name=request.querystring("username")
if name<>"" then
username=name
cz="姓名"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
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
if rs("grade")=10 and sjjh_name<>Application("sjjh_user") then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：不可以修改站长资料！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','查找：["&cz&"]="&username&"的记录！')"
Response.Write "<body text='#000000' background='../jhimg/bk_hc3w.gif' link='#0000FF' vlink='#0000FF' alink='#0000FF'>"
Response.Write "<form method=post action=updateuser.asp><table  border='1' cellspacing='0' cellpadding='5' height='20' align='center' bordercolorlight='#000000' bordercolordark='#FFFFFF' bgcolor='#FF9900' bordercolor='#FF0000'><tr><td colspan='6'><div align='center'>江湖ID:"&rs("id")&"<input type=hidden readonly size=4 name=id value="&rs("id")&">(<font color=blue size=-1>注意：数据类型与所对应的数据必需相同，否则出错！不能有空的数据存在!</font>)</div></td></tr>"
Response.Write "<tr>"
for i=1 to rs.Fields.Count-18
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
	Response.Write "<td bgcolor='#FF9900' onmouseout="&chr(34)&"this.bgColor='#FF9900';"&chr(34)&" onmouseover="&chr(34)&"this.bgColor='#8B851C';"&chr(34)&"><div align='right'><font size=-1>"&strname&"(<font color=blue size=-2>"&strtype&"</font>)：<input type=text size='10' name="&rs.Fields(i).Name&" value='"&rs.Fields(i).value&"'></font></div></td>"
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
