<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'回来♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=green>【我回来了】<font color=" & saycolor & ">"+huilai()+"</font></font>"
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'回来
function huilai()
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))=0 then
	Response.Write "<script language=JavaScript>{alert('您并没有设为暂离，请不要使用此功能！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,会员等级,身份,门派,名单头像,性别,好友名单,配偶,通缉 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
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
		sjjh_zm=sjjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&sjjh_jhdj&"|"&sjjh_id&"|"&hydj&"|0"&"|"&onlinexx(9)
		onlinelist(i)=sjjh_zm
		exit for
	end if
next
Application("sjjh_onlinelist"&nowinroom)=onlinelist
sjjh_zanli=split(application("sjjh_zanli"),";")
for x=0 to UBound(sjjh_zanli)
dxname=split(sjjh_zanli(x),"|")
if dxname(0)="!"&sjjh_name&"!" then
	dxcz=sjjh_zanli(x)&";"
	application("sjjh_zanli")=replace(application("sjjh_zanli"),dxcz,"")
	 huilai=sjjh_name&"回到自己的座位，大家有什么话想跟我说吗……"
	 exit for
end if
next
Application.UnLock
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end function
%>
<Script >
parent.crm='<%=Application("sjjh_chatroomname"&session("nowinroom"))%>';
parent.myroom=<%=session("nowinroom")%>
parent.f2.document.af.mdsx.checked=true;
parent.f2.document.af.sytemp.focus();
parent.m.location.reload();
parent.f2.location.href='../f22.asp?id=1';
</script>