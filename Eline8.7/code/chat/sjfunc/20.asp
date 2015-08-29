<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'册封♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_cf then
	Response.Write "<script language=JavaScript>{alert('提示：册封需要管理等级["&grade_cf&"]才可以操作！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if Instr(Application("sjjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('“" & towho & "”不在聊天室中，不能对其发言！');parent.f2.document.af.towho.value='大家';parent.f2.document.af.towho.text='大家';parent.m.location.reload();</script>"
	Response.end
end if
act=0
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
says="<font color=green>【册封】<font color=" & saycolor & ">"+cefen(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'册封
function cefen(fn1,to1)
fn1=trim(fn1)
if len(fn1)>10 then
	Response.Write "<script language=JavaScript>{alert(提示：身份最长不可超过10个字符！');}</script>"
	Response.End
end if
if instr(fn1,"掌门")<>0 then
	Response.Write "<script language=JavaScript>{alert('他作掌门你作什么？');}</script>"
	Response.End
end if
if len(fn1)>2 and (instr(fn1,"长老")<>0 or instr(fn1,"护法")<>0 or instr(fn1,"堂主")<>0) then
	Response.Write "<script language=JavaScript>{alert('关于:长老、护法、堂主为系统保留，请不要使用在名号中！');}</script>"
	Response.End
end if
cefeng1=instr(says,"=")
cefeng2=instr(says,",")
cefeng3=instr(says,"grade")
cefeng4=instr(says,"身份")
cefeng5=instr(says,"门派")
cefeng6=instr(says,"官府")
if cefeng1<>0 or cefeng2<>0 or cefeng3<>0 or cefeng4<>0 or cefeng5<>0 or cefeng6<>0 then
	Response.Write "<script language=JavaScript>{alert('操作错误，小子想黑我，滚！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 门派 from 用户 where 姓名='" & sjjh_name &"'" &" and 身份='掌门'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('想作什么呀，你可不是掌门！');}</script>"
	Response.End
end if
mp=rs("门派")
rs.close
rs.open "select * from 用户 where 姓名='" & to1 &"'" &" and 门派='" & mp & "'",conn,2,2
to1=rs("姓名")
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]也不是你派的弟子你想作什么？');}</script>"
	Response.End
end if
select case fn1
case "长老"
	if rs("等级")<25 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]还不够25级，不能封长老！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As 数量 from 用户 where grade=4 and 身份='长老' and 门派='"& mp &"'")
	musers=tmprs("数量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=8 then
		Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]现在你派的长老有8个了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用户 set 身份='长老',grade=4 where 姓名='" & to1 &"'"

case "护法"
	if rs("等级")<20 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]还不够20级，不能封护法！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As 数量 from 用户 where grade=3 and 身份='护法' and 门派='"& mp &"'")
	musers=tmprs("数量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=10 then
		Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]现在你派的护法有10个了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用户 set 身份='护法',grade=3 where 姓名='" & to1 &"'"

case "堂主"
	if rs("等级")<15 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]还不够15级，不能封堂主！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where grade=2 and 身份='堂主' and 门派='"& mp &"'")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
	if musers>=16 then
		Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]现在你派的堂主有16个了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用户 set 身份='堂主',grade=2 where 姓名='" & to1 &"'"
case else
	conn.execute "update 用户 set 身份='"& fn1 &"',grade=1 where 姓名='" & to1 &"'"
end select
cefen=mp&"派掌门：##册封%%为" & mp & "的<font color=red><b>" & fn1 &"</b></font>成功,大家祝贺！"
'记录
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','册封','"& cefen & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>