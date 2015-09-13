<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'心情(离席)
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<jhdj_xq then
	Response.Write "<script language=JavaScript>{alert('提示：要使用心情功能需要["&jhdj_xq&"]级以上才可以使用！');}</script>"
	Response.End
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
says=replace(says,"\","")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【心情】<font color=" & saycolor & ">"+xinqing(mid(says,i+1))+"</font>"
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'心情
function xinqing(fn1)
fn1=trim(fn1)
if len(fn1)>3 then
	Response.Write "<script language=JavaScript>{alert('心情状态长度不可大于3个字符！');}</script>"
	Response.End
end if
badword="站长,快乐,阿南,射精,奸,死,屎,妈,娘,尻,操,王八,逼,贱,狗,婊,表,靠,叉,插,干,鸡巴,睾丸,蛋,包皮,龟头,屄,赑,妣,肏,奶,尻,屌,作爱,做爱,床,抱抱,鸡八,处女,打炮,十八摸,爸,我儿,·,主席,泽民,法伦,洪志,六扇门,官府"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(fn1),bad(i))<>0 then 
		Response.Write "<script language=JavaScript>{alert('提示：输入文字不雅！');}</script>"
		Response.end
	end if
next
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,内力,操作时间,会员等级,身份,门派,名单头像,性别,好友名单,配偶,通缉,转生 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 and aqjh_grade<10 then
	s=3-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：您要使用心情功能请等["&s&"]分再操作！');}</script>"
	Response.End
end if
if rs("内力")<1000 then
	Response.Write "<script language=JavaScript>{alert('提示：需要内力1000、好好练几天吧！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用户 set 内力=内力-1000,操作时间=now() where 姓名='" & aqjh_name &"'"
aqjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
 jhsf="贵宾"
end if
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then 
 jhsf="财神"
end if
if rs("配偶")=Application("aqjh_user") and rs("性别")="女" then
 jhsf="站长夫人"
end if
if Application("aqjh_mengzhu")=aqjh_name then 
 jhsf="武林盟主"
end if
if rs("通缉")=True then
	jhmp="通缉犯"
else
	jhmp=rs("门派")
end if
jhtx=rs("名单头像")
sex=rs("性别")
myzs=rs("转生")
mypeiou=rs("配偶")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=aqjh_name then
		aqjh_zm=aqjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&aqjh_jhdj&"|"&aqjh_id&"|"&hydj&"|"&onlinexx(8)&"|"&fn1&"|"&mypeiou&"|"&myzs
		onlinelist(i)=aqjh_zm
		exit for
	end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application.UnLock
xinqing=aqjh_name&"设置自已的心情状态为["&fn1&"]……"
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end function
%>