<!--#include file="conn.asp"-->
<!--#include file="z_dgconn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="jhConst.asp"-->
<%stats="发送点歌成功"
call nav()
call head_var(0,0,"点歌控制面板","z_dglistall.asp")

dim isfull
dim mess,says,act,towhoway,towho,addwordcolor,saycolor,addsays,saystr
isfull=false		
founderr=false

if not founduser then
	errmsg=errmsg+"<br>"+"<li>客人不可以查阅点歌列表"
	errmsg=errmsg+"<br>"+"<li>请先<a href=login.asp><font color=blue>登录</font></a>,还没有<a href=reg.asp><font color=blue>注册</font></a>"
	call dvbbs_error()
else
	dim lincept,lcontent,lmedianame,lurl
	
	lincept=replace(Request.Form("incept"),"'","")
	lcontent=replace(Request.Form("content"),"'","")
	lmedianame=replace(Request.Form("medianame"),"'","")
	lurl=Request.Form("url")
	
	if lincept="" then 
		errmsg=errmsg+"<br>"+"<li>请输入接收用户名"
		founderr=true
	else
		if instr(lincept,"全体会员")>0 then
			lincept="全体会员"
			isfull=true
		else 	
			lincept=split(lincept,"|")
			if ubound(lincept)>=5 then
				errmsg=errmsg+"<br>"+"<li>最多只能发送给5个用户"
				founderr=true
			end if
		end if	
	end if
	if lmedianame="" then 
		errmsg=errmsg+"<br>"+"<li>请输入歌曲名"
		founderr=true
	end if
	if lurl="" or lurl="http://" then 
		errmsg=errmsg+"<br>"+"<li>请正确输入音乐地址"
		founderr=true
	end if
	if lcontent="" then 
		errmsg=errmsg+"<br>"+"<li>请输入祝福语"	
		founderr=true
	end if
	if len(lmedianame)>50 then '--------------判断名字
		errmsg=errmsg+"<br>"+"<li>歌曲名不能多于50字"
		founderr=true	
	end if
	if len(lcontent)>200 then '--------------判断名字
		errmsg=errmsg+"<br>"+"<li>祝福语不能多于200字"
		founderr=true	
	end if
	if isfull then
		if isnull(myvip) or myvip<>1 then
			if mymoney<clng(forum_user(18)) then
				errmsg=errmsg+"<br>"+"<li>你没有足够的现金为全体会员点歌"
				founderr=true
			end if
		else
			if mymoney<clng(forum_user(20)) then
				errmsg=errmsg+"<br>"+"<li>你没有足够的现金为全体会员点歌"
				founderr=true
			end if
		end if
	else
		if isnull(myvip) or myvip<>1 then
			if mymoney<clng(forum_user(19))*ubound(lincept) then
				errmsg=errmsg+"<br>"+"<li>你没有足够的现金为指定的会员点歌"
				founderr=true
			end if
		else
			if mymoney<clng(forum_user(21))*ubound(lincept) then
				errmsg=errmsg+"<br>"+"<li>你没有足够的现金为指定的会员点歌"
				founderr=true
			end if
		end if
	end if
	if founderr then
		call dvbbs_error()
	else
		call updata()
		call CloseDB()
		call dvbbs_suc()
	end if
end if

says="<bgsound src=wav/diangl.wav loop=1><img src=img/diangel.gif><font color=000099><b>〖点歌祝福〗</b></font>"&mess
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & jhname & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
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
call activeonline()
call footer()
'===================================
sub updata()
	dim ii
	
	if isfull then
		set rs=server.createobject("adodb.recordset")
		sql="select * from media" 
		rs.open sql,connDG,3,2
		rs.addnew
		rs("sender")=membername
		rs("incept")=lincept
		rs("sendtime")=now()
		rs("content")=lcontent
		rs("medianame")=lmedianame
		rs("url")=lurl
		rs.update
		rs.close
		set rs=nothing
		if isnull(myvip) or myvip<>1 then
			conn.execute("update [user] set userWealth=userWealth-"&forum_user(18)&" where UserName='"&membername&"'")
		else
			conn.execute("update [user] set userWealth=userWealth-"&forum_user(20)&" where UserName='"&membername&"'")
		end if
                mess="<font color=#cc0000>"&jhname&"</font><font color=006600>在江湖论坛给[<font color=navy>全体会员</font>]点了一首歌<img src=img/diange.gif>:<a href=../bbs/Z_dglistme.asp target=_blank title=点击便可查看点歌><b>{"&lmedianame&"}</b></a>，祝福道:<font color=ff0000><u>"&lcontent&"</u></font>，希望大家点击<a href=../bbs/Z_dglistme.asp target=_blank title=点击便可查看点歌><b>歌名</b></a>查看点歌！</font>"
		sucmsg="<li>点歌祝福已成功发出!<br><li>你向论坛所有的会员发送了点歌祝福"	
	else	
		for ii=0 to ubound(lincept)
			set rs=server.createobject("adodb.recordset")
			sql="select username from [user] where username='"&lincept(ii)&"'"
			rs.open sql,conn,1,1
			if rs.eof and rs.bof then
				sucmsg=sucmsg+"<br>"+"<li>论坛没有<font color=red>["&lincept(ii)&"]</font>这个用户"
			else
				rs.close
				sql="select * from media" 
				rs.open sql,connDG,3,2
				rs.addnew
				rs("sender")=membername
				rs("incept")=lincept(ii)
				rs("sendtime")=now()
				rs("content")=lcontent
				rs("medianame")=lmedianame
				rs("url")=lurl
				rs.update
				rs.close
				set rs=nothing
				dim sender,title,body
				sender=membername
				title="送给您的祝福"
				body="[color=green]"&[membername]&"[/color] 点了一首歌 [color=navy]"&lmedianame&"[/color] 给你！"&chr(10)&"祝福语：[color=blue]"&lcontent&"[/color] "&chr(10)&"[B][URL=Z_dglistme.asp]点击这里查看点歌[/URL][/B]"
				sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&lincept(ii)&"','"&sender&"','"&title&"','"&body&"',Now(),0,1)"
				conn.Execute(sql)
				if isnull(myvip) or myvip<>1 then
					conn.execute("update [user] set userWealth=userWealth-"&forum_user(19)&" where UserName='"&membername&"'")
				else
					conn.execute("update [user] set userWealth=userWealth-"&forum_user(21)&" where UserName='"&membername&"'")
				end if
                mess="<font color=#cc0000>"&jhname&"</font><font color=006600>在江湖论坛给朋友[<font color=navy>"&lincept(ii)&"</font>]点了一首歌<img src=img/diange.gif>:<a href=../bbs/Z_dglistme.asp target=_blank title=点击便可查看点歌><b>{"&lmedianame&"}</b></a>，祝福道:<font color=ff0000><u>"&lcontent&"</u></font>，希望朋友[<font color=navy>"&lincept(ii)&"</font>]点击<a href=../bbs/Z_dglistme.asp target=_blank title=点击便可查看点歌><b>歌名</b></a>查看点歌！</font>"
				sucmsg=sucmsg+"<li>点歌祝福已成功发出!系统已向<font color=green>["&lincept(ii)&"]</font>发出短消息通知"
			end if	
		next
	end if		
end sub
%>