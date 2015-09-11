<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'吸星大法
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<jhdj_xx then
	Response.Write "<script language=JavaScript>{alert('提示：吸星大法需要["&jhdj_xx&"]级才可以操作！');}</script>"
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
says="<font color=green>【吸星大法】<font color=" & saycolor & ">"+xxdf(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'吸星大法
function xxdf(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 保护,门派 from 用户 where 姓名='" & aqjh_name &"'" ,conn,2,2
if rs("门派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：出家人不可以操作!');}</script>"
	Response.End
end if
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要吸取内力!');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,保护,等级,姓名,内力,grade from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方是出家人不可以操作!');}</script>"
	Response.End
end if
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对方正在练功保护请不要偷袭!');}</script>"
	Response.End
end if
if rs("等级")<=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('请不要对初入江湖新手操作！');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你要偷管理员的钱想死呀！？！');}</script>"
	Response.End
end if
to1=rs("姓名")
nlto=rs("内力")
rs.close
randomize timer
s=int(rnd*20)+1
nlto=int((nlto/100))*s
r=int(rnd*4)
Select Case r
Case 1
	xxdf="##<bgsound src=wav/xixing.wav loop=1>使用一招吸星大法,吸取%%内力<font color=red>"& s &"%</font>,共计:<font color=red>"& nlto &"</font>点!但不小心全部丢失了!"
Case 2
	xxdf="##<bgsound src=wav/xixing.wav loop=1>想吸取%%内力,没有成功！自己失去内力<font color=red>"& s &"%</font>,共计:<font color=red>"& nlto &"</font>点!%%哈…大笑,知道我厉害了吧！"
	conn.execute "update 用户 set 内力=内力-"&nlto&" where 姓名='" & aqjh_name &"'"
case 3
	rs.open "select 银两 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
	if rs("银两")>50000 then
		conn.execute "update 用户 set 银两=银两-50000 where 姓名='" & aqjh_name &"'"
		xxdf="##<bgsound src=wav/qt.wav loop=1>偷%%的内力被发现,只好上下打点花了<font color=red>5万</font>块,逃过此劫,幸运呀!"
	else
		conn.execute "update 用户 set 状态='狱' where 姓名='" & aqjh_name &"'"
		xxdf="##<bgsound src=wav/oh_no.wav loop=1>刚刚把手搭在%%肩膀上,就被官府的人发现了,立即被捉去坐牢了"
		call boot(aqjh_name,"吸取内力，操作者："&to1&",告官让你坐牢房！")
	end if
	rs.close
case else
	xxdf="##<bgsound src=wav/xixing.wav loop=1>使用一招吸星大法,吸取%%内力<font color=red>"& s &"%</font>,共计:"& nlto &"点!气得%%吐血！"
	conn.execute "update 用户 set 内力=内力+"&nlto&" where 姓名='" & aqjh_name &"'"
	conn.execute "update 用户 set 内力=内力-"&nlto&" where 姓名='" & to1 &"'"
End Select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
