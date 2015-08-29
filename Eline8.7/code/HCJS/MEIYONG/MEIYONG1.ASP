<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


my=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.bof or rs.eof then
	response.write "你不是江湖中人，不能进美容院"
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	response.end
else
mymoney=rs("银两")
if rs("魅力")<1 or rs("银两")>100 and rs("状态")<>"狱" and rs("状态")<>"死" then
randomize timer
r=int(rnd*3)
if r=1 then
conn.execute "update 用户 set 银两=银两-100,魅力=魅力+12 where 姓名='"&sjjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "经过美容院小姐的一番摆弄,你的魅力增加了12点,美容院小姐收取了你100两银子!"
response.end
else
conn.execute "update 用户 set 银两=银两-10 where 姓名='"&sjjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "美容院的学徒由于学艺不精,弄了半天也没有帮你增加魅力.白花了你十两服务费!"
response.end
end if
else
response.write "没钱也想扮靓?美容院小姐们一番讥讽把你轰走门口"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
%>
<%end if%>
