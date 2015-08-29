<%
sub addonline(name,data,inroom)
Application.Lock
Dim newonlinelist()
js = 1
onlinelist = Application("sjjh_onlinelist" & inroom)
onlineno = UBound(onlinelist)
yjl = 0
For i = 1 To onlineno
onuser = Split(onlinelist(i), "|")
'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then      '按门派名字汉语拼音排序
'if yjl=0 and len(onuser(2))<len(jhmp) then         '按门派名字长度，长的在前
'if yjl=0 and len(onuser(2))>len(jhmp) then         '按门派名字长度,短的在前
If yjl = 0 And StrComp(onuser(0), name, 1) = 1 Then '按名字汉语拼音排序
'if yjl=0 and len(onuser(0))<len(name) then    '按名字长度，长的在前
'if yjl=0 and len(onuser(0))>len(name) then    '按名字长度,短的在前
    ReDim Preserve newonlinelist(js + 1)
    yjl = 1
    newonlinelist(js) = data
    newonlinelist(js + 1) = onlinelist(i)
    js = js + 2
Else
    ReDim Preserve newonlinelist(js)
    newonlinelist(js) = onlinelist(i)
    js = js + 1
End If
Erase onuser
Next
If yjl = 0 Then
    ReDim Preserve newonlinelist(js)
    newonlinelist(js) = data
End If
Application("sjjh_onlinelist" & inroom) = newonlinelist
Application("sjjh_useronlinename" & inroom) = Application("sjjh_useronlinename" & inroom) & name & ","
Erase newonlinelist
Erase onlinelist
Application.UnLock
end sub

sub droponline(name,inroom)
Application.Lock
onlinelist=Application("sjjh_onlinelist" & inroom)
onno=ubound(onlinelist)
for i=0 to onno
if InStr(onlinelist(i),name & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("sjjh_onlinelist" & inroom)=onlinelist
sStr=Application("sjjh_useronlinename" & inroom)
ss=instr(","& sStr,"," & name &",")
strlen=len(sStr)
Application("sjjh_useronlinename" & inroom)=mid(sStr,1,(ss-1)) & mid(sStr,(ss+len(name)+1),(strlen-ss))
Application.UnLock
end sub

function isonline(name,inroom)
if Instr(","& Application("sjjh_useronlinename" & inroom),"," & name & ",")=0 then
	isonline=false
else
	isonline=true
end if
end function

sub addsay(inroom,name,addsays,towho,addwordcolor,saycolor,says,act,towhoway)
saystr="parent.sh('" & addwordcolor & "','" & saycolor & "','" & act & "','" & name & "','" & addsays & "','" & towho & "','" & says & "','" & towhoway & "');"
Application.Lock
sd=Application("sjjh_sd" & inroom)
line=int(Application("sjjh_line" & inroom))
Application("sjjh_line" & inroom)=line+1
for i=1 to 48
 sd(i)=sd(i+2)
next
sd(49)=line+1
sd(50)=saystr
Application("sjjh_sd" & inroom)=sd
Application.UnLock
end sub
'踢出聊天室
sub boot(toname)
inroom=session("inroom")
call droponline(toname,inroom)
towhoway=0
saycolor="660099"
addwordcolor="660099"
addsays="对"
says1="<font class=t2>【江湖消息】</font>" & toname &"已经掉出聊天室……"
call addsay(inroom,sjjh_name,"对","大家","660099","660099",says1,"踢出",0)
end sub

'取聊天数据
function onlinedata(name,inroom,datasn)
Application.Lock
onlinelist=Application("sjjh_onlinelist" & inroom)
onno=ubound(onlinelist)
for i=0 to onno
if InStr(onlinelist(i),name & "|")=1 then
	if datasn=-1 then
		onlinedata=onlinelist(i)
	else
		ontemp=split(onlinelist(i),"|")
		onlinedata=ontemp(datasn)
		erase ontemp
		erase onlinelist
	end if
	exit for
end if
next
Application.UnLock
end function

'修改聊天数据
sub modionline(name,inroom,datasn,data)
Application.Lock
onlinelist=Application("sjjh_onlinelist" & inroom)
onno=ubound(onlinelist)
for i=0 to onno
if InStr(onlinelist(i),name & "|")=1 then
	ontemp=split(onlinelist(i),"|")
	ontemp(datasn)=data
	onlinelist(i)=join(ontemp,"|")
	exit for
end if
next
Application("sjjh_onlinelist" & inroom)=onlinelist
Application.UnLock
end sub
%>