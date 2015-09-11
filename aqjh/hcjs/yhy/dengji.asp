<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if aqjh_jhdj<3 then 
	Response.Write "<script Language=Javascript>alert('江湖小辈，本红园不招收无名望的姑娘！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ID,性别,魅力,配偶 FROM 用户 WHERE 姓名='" & aqjh_name & "'",conn
aqjh_id=rs("id")
sex=rs("性别")
meimao=rs("魅力")
peiou=rs("配偶")
rs.close

if sex<>"女" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('喂！有没有搞错，本红园可是只招姑娘的！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if meimao<1000 then 
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的魅力值低于1000，本园不收！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if peiou<>"无" and peiou<>"" and peiou<>"未定" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('本园不收已婚女人，麻烦！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%><!--#include file="jiu.asp"--><% 
sql="select * from 妓女 where 妓女ID="&aqjh_id
set rs=connt.execute(sql)
if rs.bof or rs.eof then
	sql="insert into 妓女(妓女ID,姓名,美貌度) values (" & aqjh_id & ",'" & aqjh_name & "'," & meimao & ")"
	connt.execute sql
	connt.close
	set connt=nothing
	conn.execute "update 用户 set 银两=银两+2000000 where 姓名='" & aqjh_name & "'"
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "恭喜，你正式成为本园的姑娘，等到卖身银两1200000万！"
	response.write "<br><a href=index.asp>返回</a>"
	response.end
end if
response.write "你已经是本红园的姑娘了，怎么还来登记呀！"
response.write "<br><a href=index.asp>返回</a>"
rs.close
set rs=nothing
conn.close
set conn=nothing
connt.close
set connt=nothing
%>