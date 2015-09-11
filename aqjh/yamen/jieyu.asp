<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name=request("name")
my=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是江湖中人，不能救人"
conn.close
response.end
else
if rs("等级")<30 then
response.write my & "江湖新手你也想劫狱？门都没有。"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("轻功")<5000 then
response.write my & "你轻功还没练到5000点就学人去劫狱？找死呀！"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("武功")>2000 and (rs("状态")<>"狱" and rs("状态")<>"牢") then
randomize timer
r=int(rnd*3)
if r=1 then
sql="update 用户 set 状态='正常' where 姓名='" & name & "'"
conn.execute sql
conn.close
Response.Redirect "laofan.asp"
else
sql="update 用户 set 状态='狱' where 姓名='" & my & "'"
conn.execute sql
conn.close
response.write "劫获不成功！你也被抓了"
response.end
end if
else
response.write "你不能劫狱"
conn.close
response.end
end if
end if
%><head></head>