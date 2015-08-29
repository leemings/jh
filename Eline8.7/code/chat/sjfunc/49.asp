<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'在线赌博♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不能够赌博！');}</script>"
	Response.End
end if
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【在线赌博】<font color=" & saycolor & ">"+dubo(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'在线赌博
function dubo(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
'进行赌场消息
if fn1="消息" then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select top 1 a FROM g WHERE c='银' and b='庄家' and f=true",conn,2,2
	if rs.eof or rs.bof then
		yinz="银局庄家：无"
	else
		yinz="银局庄家："&rs("a")
	end if
	rs.close
	rs.open "select top 1 a FROM g WHERE c='点' and b='庄家' and f=true",conn,2,2
	if rs.eof or rs.bof then
		dianz="点局庄家：无"
	else
		dianz="点局庄家："&rs("a")
	end if
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and f=true and e='大'")
	dars=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and f=true and e='小'")
	xiaors=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and f=true and e='单'")
	danrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and f=true and e='双'")
	shuangrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and f=true and d<>0")
	durs=tmprs("数量")
	dubo="消息：<img src='../jhimg/sz.gif'>[<font color=blue><b>银两局</b>,"&yinz&"</font>]押大：<font color=blue>"&dars&"</font>人,押小：<font color=blue>"&xiaors&"</font>人,押单：<font color=blue>"&danrs&"</font>人，押双：<font color=blue>"&shuangrs&"</font>人,合计：<font color=blue>"&durs&"</font>人,还差：<font color=blue>"&(10-durs)&"</font>人开局!"
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and f=true and e='大'")
	dars=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and f=true and e='小'")
	xiaors=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and f=true and e='单'")
	danrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and f=true and e='双'")
	shuangrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and f=true and d<>0")
	durs=tmprs("数量")
	dubo=dubo&"<img src='../jhimg/sz.gif'>[<font color=blue><b>存点局</b>,"&dianz&"</font>]押大：<font color=blue>"&dars&"</font>人,押小：<font color=blue>"&xiaors&"</font>人,押单：<font color=blue>"&danrs&"</font>人，押双：<font color=blue>"&shuangrs&"</font>人,合计：<font color=blue>"&durs&"</font>人,还差：<font color=blue>"&(10-durs)&"</font>人开局!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
if sjjh_jhdj<jhdj_db then
	Response.Write "<script language=JavaScript>{alert('提示：战斗不够["&jhdj_db&"]级，不可以作庄的！');}</script>"
	Response.End 
end if
if right(fn1,2)<>"作庄" or (left(fn1,1)<>"银" and left(fn1,1)<>"金" and left(fn1,1)<>"点") then 
	Response.Write "<script language=JavaScript>{alert('提示：申请作庄输入错误："&fn1&"');}</script>"
	Response.End 
end if
f=Minute(time())
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select 存款,allvalue,会员等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if left(fn1,1)="银" then	
	if rs("存款")<20000000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：[银两局]你的存款不够2000万，不能作庄！');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 a,f FROM g WHERE c='银' and b='庄家'",conn,2,2
	if rs("f")=true then
		Response.Write "<script language=JavaScript>{alert('提示：现在["&rs("a")&"]为[银局]庄家,你不能作庄！！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "update g set a='"&sjjh_name&"',f=true where c='银' and b='庄家' and f=false"
	'conn.execute "insert into g(a,b,c,d,e,f) values ('"&sjjh_name &"','庄家','银',0,'庄',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<img src='../jhimg/sz.gif'>##在江湖在线赌场[银两局]申请作庄成功，现在大家可以豪赌一把了,快来快来呀……"
end if

if left(fn1,1)="点" then
'r=Day(date())
'if r/5=int(r/5) then
'	rs.close
'	set rs=nothing	
'	conn.close
'	set conn=nothing
'	Response.Write "<script language=JavaScript>{alert('提示：存点赌博只可以在5、10、15、20、25、30号进行！');}</script>"
'	Response.End 
'end if
select case rs("会员等级")
		case 1
			hydian=31250
		case 2
			hydian=90000
		case 3
			hydian=250000
		case 4
			hydian=490000
		case else
			hydian=0
end select	
	if (rs("allvalue")-hydian)<4000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：[存点局]你的存点不够4000点，不能作庄！');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 a,f FROM g WHERE c='点' and b='庄家'",conn,2,2
	if rs("f")=true then
		Response.Write "<script language=JavaScript>{alert('提示：现在["&rs("a")&"]为[存点局]庄家,你不能作庄！！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "update g set a='"&sjjh_name&"',f=true where c='点' and b='庄家' and f=false"
'	conn.execute "insert into g(a,b,c,d,e,f) values ('"&sjjh_name &"','庄家','点',0,'庄',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<img src='../jhimg/sz.gif'>##在江湖在线赌场[存点局]申请作庄成功，现在大家可以豪赌一把了,快来快来呀……"
end if

end function
%>
