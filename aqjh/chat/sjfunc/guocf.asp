<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'国家册封
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_cf then
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
if Instr(Application("aqjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('“" & towho & "”不在聊天室中，不能对其发言。');parent.f2.document.af.towho.value='大家';parent.f2.document.af.towho.text='大家';parent.m.location.reload();</script>"
	Response.end
end if
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【国家册封】<font color=" & saycolor & ">"+cefen(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'册封
function cefen(fn1,to1)
fn1=trim(fn1)
if len(fn1)>10 then
	Response.Write "<script language=JavaScript>{alert(提示：身份最长不可超过10个字符！');}</script>"
	Response.End
end if
if instr(fn1,"君主")<>0 then
	Response.Write "<script language=JavaScript>{alert('让他作皇帝，你作什么？');}</script>"
	Response.End
end if
if len(fn1)>2 and (instr(fn1,"丞相")<>0 or instr(fn1,"将军")<>0 or instr(fn1,"侍卫")<>0) then
	Response.Write "<script language=JavaScript>{alert('关于:丞相、将军、侍卫为系统保留，请不要使用在名号中！');}</script>"
	Response.End
end if




Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 国家 from 用户 where 姓名='" & aqjh_name &"'" &" and 职位='君主'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('想作什么呀，你可不是皇帝！');}</script>"
	Response.End
end if
guojia=rs("国家")
rs.close
rs.open "select * from 用户 where 姓名='" & to1 &"'" &" and 国家='" & guojia & "'",conn,2,2
to1=rs("姓名")
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]不是你国的子民你想作什么？');}</script>"
	Response.End
end if
select case fn1
case "丞相"
	if rs("等级")<100 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]还不够100级，不能封丞相!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As 数量 from 用户 where 职位='丞相' and 国家='"& guojia &"'")
	musers=tmprs("数量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=4 then
		Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]现在你的国家已经有2个丞相了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用户 set 职位='将军'where 姓名='" & to1 &"'"

case "将军"
	if rs("等级")<80 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]还不够80级，不能封位你国的将军!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As 数量 from 用户 where 职位='将军' and 国家='"& guojia &"'")
	musers=tmprs("数量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=8 then
		Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]现在你的国家已经有4个将军了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用户 set 身份='侍卫' where 姓名='" & to1 &"'"

case ""
	if rs("等级")<60 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]还不够60级，让他做侍卫，能保护你的国家么!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 职位='侍卫' and 国家='"& guojia &"'")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
	if musers>=16 then
		Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]现在你的国家已经有8个侍卫，足够强大了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用户 set 职位='侍卫' where 姓名='" & to1 &"'"
case else
	conn.execute "update 用户 set 职位='"& fn1 &"' where 姓名='" & to1 &"'"
end select
cefen=guojia&"的皇帝下了诏书：奉天承运，皇帝诏曰，##册封%%为" & guojia & "的<font color=red><b>" & fn1 &"</b></font>成功, %%从此官运亨通。发达拉！"
'记录
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','册封','"& cefen & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>