<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%
aqjh_roominfo=split(Application("aqjh_room"),";")
nowinroom=session("nowinroom")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
onlinekill=Application("aqjh_onlinekill")
'if onlinenow<onlinekill and chatinfo(0)<>"" then
	'Response.Write "<script language=JavaScript>{alert('提示:大厅在线低于"&onlinekill&"人不得动武！');}</script>"
	'Response.End
'end if
next
%>
<%'神龙攻击
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")


if chatinfo(5)<>0 or nowinroom=3 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以发招！');}</script>"
	Response.End
end if
f=Minute(time()) 
if f>30 and nowinroom<>2 and nowinroom<>4 then
        Response.Write "<script language=JavaScript>{alert('提示：江湖PK打架时间为每小时前30分，打架请去决战！');}</script>"
	response.end
end if
if aqjh_jhdj<31 then
	Response.Write "<script language=JavaScript>{alert('提示：发招需要31级以上才可以操作！');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
   pk="【神龙攻击】"

says="<font color=red>"&pk&"<font color=" & saycolor & ">"+shen(mid(says,i+1),towho)+"</font>"
call chatsay("神龙",towhoway,towho,saycolor,addwordcolor,addsays,says)


function shen(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('你想做什么？想捣乱吗？！');}</script>"
Response.End
exit function
end if


if to1=aqjh_name or to1="大家" then
	Response.Write "<script language=JavaScript>{alert('请选择正确的动武对象！');}</script>"
       Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")

sql="select * from myanimal where username='"&aqjh_name&"' and animalname='"&trim(fn1)&"'  and rest=0"
rs.open sql,conn,1,1
      if rs.bof or rs.eof then
        Response.Write "<script language=JavaScript>{alert('您没有["&fn1&"]这种神龙!或您的神龙目前己被官府囚禁了！');}</script>"
        rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
        Response.End
       end if
gong1=rs("attack")
sm=rs("sm")
lx=rs("lei")
rs.close
''''''''''''''''''''''''''''''''''''''''''''''


sql="select 保护,门派,死亡时间 from 用户 where 姓名='" & aqjh_name&"'"
	rs.open sql,conn,1,1
       mp=rs("门派")
       baohu1=rs("保护")
        reg=rs("死亡时间")
      if baohu1=true then
	Response.Write "<script language=JavaScript>{alert('您目前正在闭关修练，少惹点江湖事好不好？！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
    '   if aqjh_grade>=6 then
    '   Response.Write "<script language=JavaScript>{alert('您是官府的，不能对江湖中人进行神龙攻击！');}</script>"
   '    rs.close
	'set rs=nothing	
	'conn.close
	'set conn=nothing
    '   Response.End
    '   end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<15 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        shen="##，你不能对%%进行攻击，他刚刚被别人杀死，还是先放过他吧！"
        exit function
end if

       rs.close


''''''''''''''''''''''''''''''''''''''''''''''''
       sql="select grade,保护,门派,体力,死亡时间,等级  from 用户 where 姓名='" & to1 & "'"
       rs.open sql,conn,1,1
       ntnt=rs("等级")
       baohu2=rs("保护")
       reg1=rs("死亡时间")
       topai=rs("门派")
       ti=rs("体力")
       tograde=rs("grade")
       if baohu2=true then
	Response.Write "<script language=JavaScript>{alert('对方正在修练中，你不能进行神龙攻击？！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if

sj=DateDiff("n",rs("死亡时间"),now())
if sj<15 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，你刚刚被别人杀死，还是明智点吧！"
        exit function
end if

       if ntnt<=30  then
       Response.Write "<script language=JavaScript>{alert('对方还是江湖新手，别这么残忍呀！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if

       if tograde>=6 then
       Response.Write "<script language=JavaScript>{alert('对方可是江湖中的官府，打他是不是不想在江湖里混了');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
       end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''
       conn.execute ("update 用户 set 体力=体力-"&gong1&" where 姓名='"&to1&"'")
       conn.execute ("update myanimal set attack=attack*9/10 where username='"&nickname&"' and animalname='"&trim(fn1)&"'")

			if ti-gong1<-1000 then
                    	       conn.execute"update 用户 set 状态='死' where 姓名='" & to1 & "'"
				conn.execute"update 用户 set allvalue=allvalue+"  & ntnt*5 & ",道德=道德-10,魅力=魅力-10,杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & aqjh_name & "'"
				conn.execute("update myanimal set username='无',rest=1 where username='" &towho& "'")
call boot(to1,"神龙，操作者："&aqjh_name&",["&menpai&"]"&fn1)
                           shen="["&aqjh_name&"]放出护身神龙<b><font color=black>"&fn1&"("&lx&")</font></b>,<b><font color=red>"&sm&"</font></b><bgsound src=READONLY/long.mid loop=1>，杀伤["&to1&"]体力<b><font color=red>"&gong1&"</font></b>,"&to1&"学艺不精，当场被咬死！"&aqjh_name&"为杀这个人自己损失了10点道德和10点魅力,得到"  & ntnt*5 & "个经验，"&aqjh_name&"的神龙攻击力下降1/10"
                             rs.close
                             set rs=nothing	
	                      conn.close
	                      set conn=nothing	
                             exit function
                    end if
                       shen="["&aqjh_name&"]放出护身神龙<b><font color=black>"&fn1&"("&lx&")</font></b>,<b><font color=red>"&sm&"</font></b><bgsound src=READONLY/long.mid loop=1>，杀伤["&to1&"]体力<b><font color=red>"&gong1&"</font></b>,"&aqjh_name&"的神龙攻击力下降1/10!" 
                        rs.close
                        set rs=nothing	
	                 conn.close
                       set conn=nothing  
                           
        
end function
%>
</body>
</html>