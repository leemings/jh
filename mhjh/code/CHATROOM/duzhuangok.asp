<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
un=session("yx8_mhjh_username")
grade=session("yx8_mhjh_usergrade")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if un="" then Response.Redirect "../error.asp?id=016"
if grade<10 then
Response.Write "<script language=javascript>{alert('赌博需要10级以上\n按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chatroom")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if	
mg=abs(int(clng(Request.form("money"))))
idd=Request.form("id")
abcc=mid(mg,6)
for t=1 to len(abcc)
abc=mid(abcc,t,1)
if abc<>"0" and abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('["&abcc&"]操作错误，数量请使用数字！');parent.history.go(-1);}</script>"
	Response.End 
end if
next
yacz=idd
yinliang=abs(mg)
select case idd
	case "大"
		duboimg="<img src=../chatroom/image/da.gif>"
	case "小"
		duboimg="<img src=../chatroom/image/xiao.gif>"
	case "单"
		duboimg="<img src=../chatroom/image/dan.gif>"
	case "双"
		duboimg="<img src=../chatroom/image/shuang.gif>"
case else
	Response.Write "<script language=JavaScript>{alert('押操作为：大、小、单、双！');parent.history.go(-1);}</script>"
	Response.End 
end select
	if yinliang<100000 or yinliang>5000000 then
		Response.Write "<script language=JavaScript>{alert('最少押：10万两，最多500万两！');parent.history.go(-1);}</script>"
		Response.End 
	end if
'开始判断
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 银两 FROM 用户 WHERE 姓名='" & un &"'"
set rs=conn.execute(sql)
if rs("银两")<yinliang then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你钱好象不够！');parent.history.go(-1);}</script>"
	Response.End
end if
	sql= "Select count(*) As 数量 from d where yl<>0 and c='银'"
	set rs=conn.execute(sql)
	durs=rs("数量")
	if durs>=10 then
	        rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('["&durs&"]现在开局，稍后再下注！');parent.history.go(-1);}</script>"
		Response.End 		
	end if
	sql="select top 1 xm FROM d WHERE sf='庄家'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('现在没有庄家，不能在线赌博！如果您条件符合可以抢庄！');parent.history.go(-1);}</script>"
		Response.End 		
	end if	
	sql="select top 1 * FROM d WHERE xm='"& un &"' and c='银'"
	set rs=conn.execute(sql)
	if not(rs.eof) or not(rs.bof) then
		if rs("sf")="庄家" then
			temp=""&un&"你现在是庄呀，你要搞什么呀!"
		else
			temp=""&un&"你压["&rs("ydx")&"]"&rs("yl")&"两等着开吧！"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');parent.history.go(-1);}</script>"
		Response.End 
	end if
	conn.execute "update 用户 set 银两=银两-"&yinliang&" where 姓名='"&un&"'"
	conn.execute "insert into d(xm,sf,yl,ydx,sj,c) values ('"&un&"','玩家',"&yinliang&",'"&yacz&"',now(),'银')"	
	tmprs=conn.execute("Select count(*) As 数量 from d where ydx='大' and c='银'")
	dars=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from d where ydx='小' and c='银'")
	xiaors=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from d where ydx='单' and c='银'")
	danrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from d where ydx='双' and c='银'")
	shuangrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from d where yl<>0 and c='银'")
	durs=tmprs("数量")


'开局了
if durs>=5 then
	Randomize
	m1 = Int(6 * Rnd + 1)
	Randomize
	m2 = Int(6 * Rnd + 1)
	Randomize
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'查找庄家
sql= "select * FROM d WHERE sf='庄家' and c='银'"
set rs=conn.execute(sql)
zhuangjia=rs("xm")
rs.close
set rs=nothing
'豹子处理
if m1=m2 and m2=m3 and m3=m1 then
	sql="select yl FROM d WHERE yl<>0 and sf<>'庄家' and c='银'"
	set rs=conn.execute(sql)
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("yl")
	rs.movenext
	loop
	rs.close
	newyingyin=int(yingyin*0.9)
	conn.execute "update 用户 set 银两=银两+"&newyingyin&"+500000 where 姓名='"& zhuangjia &"'"
	conn.execute "delete * from d where  xm<>'' and c='银'"
	says="<b><font color=FF0000>【在线赌博】</b></font>庄家开：<img src=../chatroom/image/"&m1&".gif><img src=../chatroom/image/"&m2&".gif><img src=../chatroom/image/"&m3&".gif>计："& sjdubo &"点!庄家开出豹子……通杀！庄家：<font color=FF0000>["&zhuangjia&"]</font>收入："&yingyin&"两！胜者扣10%的官税。"
	call chat(says)
	Response.Write "<script language=JavaScript>{alert('[操作完成,开赌局]');parent.history.go(-1);}</script>"	
	Response.End 
end if

'初始化数据
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'开单双处理
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src=../chatroom/image/shuang.gif>"
	sql="select yl,xm FROM d WHERE ydx='双' and c='银'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("yl")
		shuangname=shuangname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update 用户 set 银两=银两+"&newyingyin&" where 姓名='"& rs("xm") &"'"
		conn.execute "update 用户 set 银两=银两-"& rs("yl") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from d where ydx='双'"
else
	danshuang="<img src=../chatroom/image/dan.gif>"
	sql="select yl,xm FROM d WHERE ydx='单' and c='银'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("yl")
		danname=danname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update 用户 set 银两=银两+"&newyingyin&" where 姓名='"& rs("xm") &"'"
		conn.execute "update 用户 set 银两=银两-"& rs("yl") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from d where ydx='单'"
end if

'对开大小处理
if sjdubo<=10 then
	daxiao="<img src=../chatroom/image/xiao.gif>"
	sql="select yl,xm FROM d WHERE ydx='小' and c='银'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("yl")
		xiaoname=xiaoname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update 用户 set 银两=银两+"&newyingyin&" where 姓名='"& rs("xm") &"'"
		conn.execute "update 用户 set 银两=银两-"& rs("yl") &" where 姓名='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "delete * from d where  ydx='小'"
else
	daxiao="<img src=../chatroom/image/da.gif>"
	sql="select yl,xm FROM d WHERE ydx='大' and c='银'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("yl")
		daname=daname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update 用户 set 银两=银两+"&newyingyin&" where 姓名='"& rs("xm") &"'"
		conn.execute "update 用户 set 银两=银两-"& rs("yl") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from d where  ydx='大'"
end if
'对剩下输的用户银两
	rs.open "select yl,xm FROM d WHERE yl<>0 and sf<>'庄家' and c='银'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("yl")
	rs.movenext
	loop
	rs.close
	newyingyin=int(yingyin*0.9)+500000
	conn.execute "update 用户 set 银两=银两+"&newyingyin&" where 姓名='"& zhuangjia &"'"
	conn.execute "delete * from d where  xm<>'' and c='银'"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	duboname=shuangname&danname&xiaoname&daname
	says="<b><font color=FF0000>【在线赌博】</b></font>庄家开：<img src=../chatroom/image/"&m1&".gif><img src=../chatroom/image/"&m2&".gif><img src=../chatroom/image/"&m3&".gif>计："& sjdubo &"点!"&danshuang&daxiao&"总下注："&zong&"两，庄家：["&zhuangjia&"]收入："&yingyin &"两,赔出："&pei&"两，合计："&yingyin-pei&"两,共有：<font color=red>"&duboname&"</font>玩家幸运押中！胜者扣10%官税!"
    call chat(says)
    Response.Write "<script language=JavaScript>{alert('[操作完成,开赌局]');parent.history.go(-1);}</script>"	
    Response.End 
    end if
    says="<b><font color=FF0000>【在线赌博】</b>"&un&"</font>从自己的小荷包里拿出:"&yinliang&"两，我买"&duboimg&"！，一定中的！现在下注如下：押大："& dars &"个，押小:"& xiaors &"个,押单："&danrs&"个,押双:"&shuangrs&"个！还差:"&(5-durs)&"个开局！"
	call chat(says)
	Response.Write "<script language=JavaScript>{alert('[下注成功]');parent.history.go(-1);}</script>"	
    Response.End 
	
sub chat(says)
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr	
end sub
%>

