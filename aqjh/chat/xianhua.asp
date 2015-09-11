<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<SCRIPT LANGUAGE="JavaScript">
if(window.name!="d"){
var i=1;while(i<=50){
window.alert("刷钱是吗？喜欢是吗？点啊，刷啊！！");
i=i+1;}top.location.href="../exit.asp"
}
</script>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
fromname=LCase(trim(request.querystring("fromname")))
yn=LCase(trim(request.querystring("yn")))
toname=LCase(trim(request.querystring("toname")))
huaming=LCase(trim(request.querystring("huaming")))
xhsl=CInt(trim(request.querystring("wpsl")))
if InStr(fromname,"or")<>0 or InStr(fromname,"'")<>0 or InStr(fromname,"`")<>0 or InStr(fromname,"=")<>0 or InStr(fromname,"-")<>0 or InStr(fromname,",")<>0 then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
if toname<>aqjh_name then
	if fromname<>aqjh_name then
		Response.Write "<script language=JavaScript>{alert('你想作什么，人家"&fromname&"也不是送给你的！');}</script>"
		Response.End
	else
		Response.Write "<script language=JavaScript>{alert('有病呀"&fromname&"？你想作什么?送给别人自己去收，真是无聊！');}</script>"
		Response.End
	end if
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 性别,配偶,情人,w7 from 用户 where 姓名='"&fromname&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close	
	set rs=nothing	
	conn.close	
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('送花的人不存在，可能有作弊行为，请跟站长反映！');}</script>"	'测4
	Response.End	
end if
if  mywpsl(rs("w7"),huaming)<xhsl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["&fromname&"]的物品数量不够("&xhsl&")个！');}</script>"
	Response.End
end if
myhua=abate(rs("w7"),huaming,xhsl)
fromsex=rs("性别")
peiou=rs("配偶")
qingren=rs("情人")
conn.execute "update 用户 set w7='"&myhua&"' where  姓名='"&fromname&"'"
rs.close
rs.open "select 性别 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
tosex=rs("性别")
randomize timer
r=int(rnd*3)
if yn=1 then
	if fromsex=tosex then
		if fromsex="男" then
			conn.execute "update 用户 set 道德=道德-500 where 姓名='"&aqjh_name&"'"
			conn.execute "update 用户 set 道德=道德-250 where 姓名='"&fromname&"'"
			Response.Write "<script language=JavaScript>{alert('你想作什么?想跟人搞玻璃关系呀？太不道德了！');}</script>"
			xianhua="最新报导：["& aqjh_name &"]接收了{"& fromname &"}送的"& huaming &"，由于两人有搞玻璃关系倾向被人发现，"& aqjh_name &"道德下降500点，"& fromname &"道德下降250点!<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]想跟{"& fromname &"}搞玻璃关系被人发现，到处发出一片谴责声，唉！不道德行为，大家小心了。。。</marquee>"
		else
			conn.execute "update 用户 set 魅力=魅力+200 where 姓名='"&aqjh_name&"'"
			conn.execute "update 用户 set 魅力=魅力+100 where 姓名='"&fromname&"'"
			xianhua="恭喜：["& aqjh_name &"]高兴地接收了{"& fromname &"}送的"& huaming &"，"& aqjh_name &"魅力上升200点，"& fromname &"魅力上升100。<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了{"& fromname &"}送的鲜花，大家觉得这两个女孩子都好有个性，两人的魅力都上升了。。。</marquee>"
		end if
	else
		conn.execute "update 用户 set 魅力=魅力+"& xhsl*100 &" where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+"& xhsl*500 &" where 姓名='"&fromname&"'"
		xianhua="恭喜：["& aqjh_name &"]高兴地接受了{"& fromname &"}送的"& huaming &"，"& aqjh_name &"魅力上升"& 100*xhsl &"点，"& fromname &"魅力上升"& 500*xhsl &"，别的大虾看得很是眼红，嘻嘻。。。"
		if peiou=aqjh_name then
			select case r
				case 0
					if tosex="女" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了老公{"& fromname &"}送的鲜花，满脸一付幸福模样，好可爱呀！~~~~这年头送老婆鲜花的人可不多了，大家觉得{"& fromname &"}可真是个模范丈夫！</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了老婆{"& fromname &"}送的鲜花，满脸一付幸福模样，今晚大概又睡不着了！~~~~这年头送老公鲜花的人可不多了，大家觉得{"& fromname &"}可真是个好妻子！</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了{"& fromname &"}送的鲜花，禁不住深情地对{"& fromname &"}说：我爱死你了！下辈子我们还要做夫妻！</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了老公{"& fromname &"}送的鲜花，高兴死了，不由得抱住{"& fromname &"}狂啵儿，真是的......哈，儿童不宜！</marquee>"
			end select
		elseif qingren=aqjh_name then
			select case r
				case 0
					if tosex="女" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了老情人{"& fromname &"}送的鲜花，满脸一付幸福模样，好可爱呀！~~~~大家觉得{"& fromname &"}可真是有点那个。。。！</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了小情人{"& fromname &"}送的鲜花，满脸一付陶醉模样，今晚大概又睡不着了！~~~~这年头，这样子的也有了，呵呵~~~~</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了情人{"& fromname &"}送的鲜花，禁不住深情地对{"& fromname &"}说：我爱死你了！下辈子我们要做夫妻！</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了情人{"& fromname &"}送的鲜花，高兴死了，两人不由得抱住狂啵儿，真是的......哈，也不怕对方的丈夫/妻子看到，真是好过份耶！</marquee>"
			end select
		else
			select case r
				case 0
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了{"& fromname &"}送的鲜花，兴奋过度，差点晕倒在地！嘿，真是罪过。。。阿弥陀佛</marquee>"
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了{"& fromname &"}送的鲜花，{"& fromname &"}兴奋得满脸通红，两眼发光，心想：“这下我有戏了！”</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]收了{"& fromname &"}送的鲜花，["& aqjh_name &"]高兴异常，不由得向{"& fromname &"}抛了几个媚眼，真是的......哈哈，也不害羞</marquee>"
			end select
		end if
	end if
else
	if fromsex=tosex then
		if fromsex="男" then
			conn.execute "update 用户 set 道德=道德+100 where 姓名='"&aqjh_name&"'"
			conn.execute "update 用户 set 道德=道德-250,魅力=魅力-1000 where 姓名='"&fromname&"'"
			xianhua="【最新报导】：["& aqjh_name &"]拒绝接收{"& fromname &"}送的"& huaming &"，"& aqjh_name &"道德上升100点，"& fromname &"道德下降250点!<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>{"& fromname &"}想跟["& aqjh_name &"]搞玻璃关系被拒绝，["& aqjh_name &"]生气地对{"& fromname &"}说道：“唉！你这种想法是不道德行为，走开。。。”</marquee>"
		else
			conn.execute "update 用户 set 魅力=魅力+100 where 姓名='"&aqjh_name&"'"
			conn.execute "update 用户 set 魅力=魅力+50 where 姓名='"&fromname&"'"
			xianhua="【最新报导】：["& aqjh_name &"]拒绝接收{"& fromname &"}送的"& huaming &"，大家觉得这两个女孩子都蛮有意思，"& aqjh_name &"魅力上升100点，"& fromname &"魅力上升50。<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了{"& fromname &"}送的鲜花，哈，有意思，没见过女孩子给女孩子送花。。。</marquee>"
		end if
	else
		conn.execute "update 用户 set 魅力=魅力+100 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力-1000 where 姓名='"&fromname&"'"
		xianhua="【最新报导】：["& aqjh_name &"]拒绝接收{"& fromname &"}送的"& huaming &"，说道：真不好意思，我不想接受你的鲜花，请你送给爱你的人吧……<font color=red>"& aqjh_name &"魅力上升100点，"& fromname &"魅力降底1000 ！</font>"
		if peiou=aqjh_name then
			select case r
				case 0
					if tosex="女" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了老公{"& fromname &"}送的鲜花，满脸怒容地说道：“你这死东西，背着我在外面胡搞，现在才想来讨好我，连窗都没有！”大家觉得{"& fromname &"}可真是有点太那个......不应该！</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了老婆{"& fromname &"}送的鲜花，满脸憔悴地说道：“我们还是离了好，这日子没法过了......”大家用怀疑的眼光看着{"& fromname &"}，心想：莫非这女人在外面......？</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了另一半{"& fromname &"}送的鲜花，有点心痛地对{"& fromname &"}说道：“不用这样了，我们还是先分开一段时间，给对方一个自由的空间吧，好考虑清楚以后要怎么过......”</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了另一半{"& fromname &"}送的鲜花，有点无所谓地对{"& fromname &"}说道：“我们已经无可挽回了，你也不必浪费这个钱了。”大家为这对曾经的模范好夫妻成为今天这样子感到痛心......</marquee>"
			end select
		elseif qingren=aqjh_name then
			select case r
				case 0
					if tosex="女" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了老情人{"& fromname &"}送的鲜花，笑着道：“你这死东西，快回头看看，你那口子找你来了。”大家不由得幸灾乐祸地看着{"& fromname &"}：“哈哈哈......背着老婆在外面胡搞，这下要回去跪搓衣板了~~”</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了小情人{"& fromname &"}送的鲜花，满脸憔悴地说道：“我们还是分手吧......我那口子盯得太紧了，唉...”大家用可怜地看着{"& fromname &"}，心想：这家伙看来天天在家里当床头柜，哈哈，真有种，也敢在外面找情人</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了情人{"& fromname &"}送的鲜花，满脸怒容地说道：“你这死人，好几天不理我，现在又想来讨好我，连窗都没有！”大家觉得{"& fromname &"}可真是好笑，活该！哈哈哈~~</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了情人{"& fromname &"}送的鲜花，笑嘻嘻地说道：“是不是又找了一个？这么长时间没来看我，现在我不想理你了......（我也想去再找一个）”</marquee>"
			end select
		else
			select case r
				case 0
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了{"& fromname &"}送的鲜花，大义凛然地说道：“别来找我，我早就已经名花有主了。”</marquee>"
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了{"& fromname &"}送的鲜花，笑嘻嘻地说道：“你想干嘛呀？是不是有什么阴谋？我可是怕怕的哟......”</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>【小道消息】</b></font>["& aqjh_name &"]拒绝了{"& fromname &"}送的鲜花，一脸不屑地看着{"& fromname &"}道：“找情人找到我头上了？你可真是晕了头，瞎了眼！”</marquee>"
			end select
		end if
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
says="<font color=red><b>【江湖消息】</b></font>"&xianhua	
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>