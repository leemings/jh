<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const3.asp"-->
<%Response.Buffer=true
Response.Expires = 0
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatroomsn=session("nowinroom")
mychatroomsn=Request.QueryString("roomsn")
chatroomname=Request.QueryString("chatroomname")
chatroomname=Application("aqjh_chatroomname"&mychatroomsn)
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
if chatroomsn=mychatroomsn then
  response.write "<Script language=javascript>alert('你已经在["&chatroomname&"]不能重复进入!');parent.m.location.reload();parent.r.location.reload();</script>"
  response.end
end if
chatroominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(chatroominfo)-1
i=mychatroomsn
online=split(trim(Application("aqjh_useronlinename"&i))," ")
onlinenum=ubound(online)+1
sj_chat_info=split(chatroominfo(i),"|")
num=int(sj_chat_info(1))
if onlinenum>=num then
  response.write "<Script>alert('["&chatroomname&"]房间当前人数已满，不能进入！');parent.r.location.reload();</script>"
  response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 and chatroomname<>"泡点专厢" and aqjh_grade<10 then
   response.write "<Script>alert('提示:泡点中,无法换房!');parent.r.location.reload();</script>"
   response.end
end if
'决战房的判断
if application("aqjh_user")<>aqjh_name then
if chatroomname=aqjh_chat2 and (Weekday(date())<>7 or Hour(time())<>21 or (Hour(time())=21 and Minute(time())>50)) then
	response.write "<Script>alert('决战爱情只有在每周六的21:00-21:50分才可进入 \n现在时间"&time()&"！时间未到！请稍等！');parent.r.location.reload();</script>"
	response.end
end if
if chatroomname=aqjh_chat2 and aqjh_grade>5 then
	response.write "<Script>alert('提示：官府中人不得参与！');parent.r.location.reload();</script>"
	response.end
end if
end if
if chatroomname=aqjh_chat2 and Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then
	response.write "<Script>alert('提示：你都是财神了，就不再争夺了！');parent.r.location.reload();</script>"
	response.end
end if
'决战房的判断END
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,性别,门派,身份,名单头像,会员等级,操作时间,通缉,转生,配偶 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
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
hydj=rs("会员等级")
myzs=rs("转生")
mypeiou=rs("配偶")
sj=DateDiff("n",rs("操作时间"),now())
if sj<1 and aqjh_grade<6 then
	s=1-sj
	Response.Write "<script language=JavaScript>{alert('提示：转换房间请等["&s&"分钟]再操作！');parent.m.location.reload();parent.r.location.reload();}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if cstr(sj_chat_info(2))=1 and aqjh_grade<7 then
	rs.close
	rs.open "select id from 用户 where 姓名='" & aqjh_name &"'"&" and "&sj_chat_info(4),conn
	if rs.eof or rs.bof  then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><meta http-equiv='pragma' content='no-cache'></head><body bgcolor=" + Application("aqjh_chatbgcolor") + " background=" + Application("aqjh_chatimage") + " bgproperties=fixed>"
		Response.Write "<script language=JavaScript>{alert('进入["&sj_chat_info(0)&"]的条件是："&sj_chat_info(3)&"');parent.r.location.reload();}</script>"
		Response.End
	end if
	rs.close
end if
'姓名,性别,门派,身份,头象,等级,id
myzanli=0
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 then myzanli=1
myonline = aqjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & aqjh_jhdj& "|" & aqjh_id& "|" & hydj&"|"&myzanli&"|"&"正常"&"|"&mypeiou&"|"&myzs
if sj_chat_info(7)<>0 or jhmp="天网" then
	conn.Execute "update 用户 set 保护=false,操作时间=now() where 姓名='" & aqjh_name &"'"	
else
	conn.Execute "update 用户 set 保护=true,操作时间=now() where 姓名='" & aqjh_name &"'"	
end if
set rs=nothing	
conn.close
set conn=nothing
'退出原房间
Application.Lock
mychatroomname=Application("aqjh_chatroomname"&chatroomsn)
onlinelist=Application("aqjh_onlinelist"&chatroomsn)
onno=ubound(onlinelist)
for i=1 to onno
 if InStr(onlinelist(i),aqjh_name & "|")=1 then
'myonline=onlinelist(i)
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  Application("aqjh_onlinelist"&chatroomsn)=onlinelist
  exit for
 end if
next
Application("aqjh_useronlinename"&chatroomsn)=Replace(Application("aqjh_useronlinename"&chatroomsn)," " & aqjh_name & " ","")
'进入新房间
dim newonlinelist()
js=1
onlinelist=Application("aqjh_onlinelist"&mychatroomsn)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")
'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'按门派名字汉语拼音排序
'if yjl=0 and len(onuser(2))<len(jhmp) then			'按门派名字长度，长的在前
'if yjl=0 and len(onuser(2))>len(jhmp) then			'按门派名字长度,短的在前
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'按名字汉语拼音排序
'if yjl=0 and len(onuser(0))<len(aqjh_name) then	'按名字长度，长的在前
'if yjl=0 and len(onuser(0))>len(aqjh_name) then	'按名字长度,短的在前
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0  then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if
Application("aqjh_onlinelist"&mychatroomsn)=newonlinelist
Application("aqjh_useronlinename"&mychatroomsn)=Application("aqjh_useronlinename"&mychatroomsn)& " "&aqjh_name & " "
erase  newonlinelist
erase  onlinelist
' onlinelist=Application("aqjh_onlinelist"&mychatroomsn)
' onlineno=ubound(onlinelist)+1
' ReDim Preserve onlinelist(onlineno)
'onlinelist(onlineno)=myonline
'Application("aqjh_onlinelist"&mychatroomsn)=onlinelist
'Application("aqjh_useronlinename"&mychatroomsn)=Application("aqjh_useronlinename"&mychatroomsn)& " "&aqjh_name & " "
session("nowinroom")=mychatroomsn
Application.UnLock	

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
act="退出"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
says="<bgsound src=wav/folder.wav loop=1><font color=black>【换房】</font><font color=#009933><font color=red>" & aqjh_name & "</font>施展出“凌波微步”轻功，转眼间便从〖" & mychatroomname & "〗消失了，原来是去【" &chatroomname & "】了。</font><font class=t>(" & time() & ")</font>"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & chatroomsn & ");<"&"/script>"
addmsg saystr
'session("SayCount")=Application("SayCount")
act="进入"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
says="<bgsound src=wav/folder.wav loop=1><font color=black>【换房】</font><font color=#009933><a href=javascript:parent.sw('[" & aqjh_name & "]'); target=f2>" & aqjh_name & "</a>施展出“凌波微步”轻功，转眼间从〖" & mychatroomname & "〗来到了【" &chatroomname & "】。</font><font class=t>(" & t & ")</font><bgsound src='readonly/cd.mid' loop='1'>"	'受话者
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & session("nowinroom") & ");<"&"/script>"
addmsg saystr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
%>
<Script >
parent.crm='<%=Application("aqjh_chatroomname"&session("nowinroom"))%>';
parent.myroom=<%=session("nowinroom")%>
parent.f2.document.af.mdsx.checked=true;
parent.f2.document.af.sytemp.focus();
parent.f2.document.af.towho.value="大家";
parent.m.location.reload();
//parent.f2.location.href='f2.asp?id=1';
parent.f2.location.reload();
parent.qp();
alert('欢迎来到〖<%=chatroomname%>〗！');
</script>