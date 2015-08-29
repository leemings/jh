<%@ LANGUAGE=VBScript codepage ="936" %>

<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'寻找秘笈
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【寻找秘笈】<font color=" & saycolor & ">"+seek(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'寻找秘笈
function seek(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,轻功,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*5)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*5)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
dj=rs("等级")

light=rs("轻功")
if rs("轻功")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的轻功不够无法施展呀，至少也得50000点啊！');window.close();}</script>"
	response.end
end if
if dj<20 then
Response.Write "<script language=JavaScript>{alert('此功能需要20级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

rs.close
conn.execute "update 用户 set 轻功=轻功-50000,操作时间=now()  where 姓名='"&sjjh_name&"'"
randomize 
r=int(rnd*9)+1
select case r
case 1
seek=sjjh_name & "你在武林中到处寻找轻功秘笈,突然在孔明家中发现一本消失已久的<img src='img/02.jpg' width='30' height='25'><font color=red>【佛山无影】</font>立刻偷走这本秘笈回家练习."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"佛山无影",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close


case 2
seek=sjjh_name & "你在武林中到处寻找轻功秘笈,但迟迟未寻获到秘笈的下落突然" & to1 & "跑来告诉你他有一本失传慎久的<img src='img/02.jpg' width='30' height='25'><font color=red>【醉佛无影】</font>" & to1 & "问道说你想要吗这本送给你."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"醉佛无影",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

	
case 3
seek=sjjh_name & "你在武林中不断到处寻找轻功秘笈,但没想到你爱人手上早有一本<img src='img/02.jpg' width='30' height='25'><font color=red>【飞燕投林】</font>他就是怕你学会了轻功到处跑就不要他了.但" & sjjh_name & "的爱人最后还是把秘笈给了" & sjjh_name & "."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"飞燕投林",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 4
seek=sjjh_name & "你在武林中不断到处寻找轻功秘笈,但想都想不到让你在南海岛寻到<img src='img/02.jpg' width='30' height='25'><font color=red>【凌波微步】</font>这样都能被你找，真是佩服."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"凌波微步",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
	
case 5
seek=sjjh_name & "你在武林中不断到处寻找轻功秘笈,有幸让你路过少林寺碰到幻尘大师他把<img src='img/02.jpg' width='30' height='25'><font color=red>【佛光化体】</font>让你带回家学习.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"佛光化体",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 6
seek=sjjh_name & "好可惜，你寻遍了江湖各个角落也没有找到什么轻功秘笈!" 
	conn.execute "update 用户 set 轻功=轻功-100 where 姓名='" & sjjh_name &"'"

case 7
seek=sjjh_name & "真幸运，你寻遍了江湖各个角落没有找到什么轻功秘笈，轻功增加了100点!" 
	conn.execute "update 用户 set 轻功=轻功+500 where 姓名='" & sjjh_name &"'"
	
case 8
seek=sjjh_name & "<bgsound src=wav/m2.mid loop=1>你在武林中不断到处寻找轻功秘笈,有天路过伟小宝家门口看见伟小宝.伟小宝说道:老子老婆太多没空学神行百变你喜欢拿去吧<img src='img/02.jpg' width='30' height='25'>" & sjjh_name & "真是幸运."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"神行百变",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close


case 9
seek=sjjh_name & "<bgsound src=wav/m2.mid loop=1>你在武林中不断到处寻找轻功秘笈,有天路过古墓派让你遇到小龙女有幸他把武林中最厉害的轻功<img src='img/02.jpg' width='30' height='25'><font color=red>【独步天下】</font>传授了给你.真是有福之人.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"独步天下",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

case 10
seek=sjjh_name & "<bgsound src=wav/m2.mid loop=1>你在武林中不断到处寻找轻功秘笈,有天路过古墓派让你遇到小龙女有幸他把武林中最厉害的轻功<font color=red>【独步天下】</font>传授了给你.真是有福之人.."
	rs.open "SELECT w10 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
	duyao=add(rs("w10"),"独步天下",1)
	conn.execute "update 用户 set  w10='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close

	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
