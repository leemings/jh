<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'心情♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_jhdj<jhdj_xq then
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
says=Replace(says,chr(34),"")
says=Replace(says,"|","")
says=Replace(says,"'","")
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
badword="站长,阿男,阿南,射精,奸,死,屎,妈,娘,尻,操,王八,逼,贱,狗,婊,表,靠,叉,插,干,鸡巴,睾丸,蛋,包皮,龟头,屄,赑,妣,肏,奶,尻,屌,作爱,做爱,床,抱抱,鸡八,处女,打炮,十八摸,爸,我儿,·,主席,泽民,法伦,洪志,六扇门,官府"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(fn1),bad(i))<>0 then 
		Response.Write "<script language=JavaScript>{alert('提示：输入文字不雅！');}</script>"
		Response.end
	end if
next
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,内力,操作时间,会员等级,身份,门派,名单头像,性别,好友名单,配偶,通缉 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 and sjjh_grade<10 then
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
conn.execute "update 用户 set 内力=内力-1000,操作时间=now() where 姓名='" & sjjh_name &"'"
sjjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
if rs("通缉")=True then
	jhmp="通缉犯"
else
	jhmp=rs("门派")
end if
jhtx=rs("名单头像")
sex=rs("性别")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("sjjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=sjjh_name then
		sjjh_zm=sjjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&sjjh_jhdj&"|"&sjjh_id&"|"&hydj&"|"&onlinexx(8)&"|"&fn1
		onlinelist(i)=sjjh_zm
		exit for
	end if
next
Application("sjjh_onlinelist"&nowinroom)=onlinelist
Application.UnLock
xinqing=sjjh_name&"设置自已的心情状态为["&fn1&"]……"
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end function
%>