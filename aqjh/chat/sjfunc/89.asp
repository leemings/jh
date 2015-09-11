<!--#include file="sjfunc.asp"-->
<%'屏蔽
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【屏蔽】<font color=" & saycolor & ">"+pingbi(towho,mid(says,i+1))+"</font>"
'towhoway=1
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'好友
function pingbi(to1,fn1)
fn1=trim(fn1)
select case fn1
case "加入"
 if Instr(LCase(application("qzld_pingbi")),LCase(";"&aqjh_name&"|"&to1&";"))>0 then
   Response.Write "<script Language=javascript>alert('您已经把人家屏蔽了，不必再屏蔽人家了！');</script>"
  response.end
' elseif Instr(LCase(application("qzld_pingbi")),LCase(";"&to1&"|"&aqjh_name&";"))>0 then
'   Response.Write "<script Language=javascript>alert(''''你的说话对象已经把你屏蔽了，你不能再对人家说话了！'''');</script>"
'  response.end
 end if
 Application.Lock
  if Application("qzld_pingbi")="" then
   Application("qzld_pingbi")=";"&aqjh_name&"|"&to1&";"
  else
   Application("qzld_pingbi")=Application("qzld_pingbi")&aqjh_name&"|"&to1&";"
  end if
 Application.UnLock
 pingbi="##已经成功地把["&to1&"]加入屏蔽名单，["&to1&"]不能再对{"&aqjh_name&"}发言了!<font class=t>("&time&")</font>"

case "删除"
 if Instr(LCase(application("qzld_pingbi")),LCase(";"&aqjh_name&"|"&to1&";"))=0 then
   Response.Write "<script Language=javascript>alert('说话对象不在你的屏蔽名单中，不必这样操作！');</script>"
  response.end
 end if
 Application.Lock
   Application("qzld_pingbi")=Replace(Application("qzld_pingbi"),";"&aqjh_name&"|"&to1&";",";")
 Application.UnLock
 pingbi="##已经成功地把["&to1&"]从屏蔽名单中删除，["&to1&"]可以对{"&aqjh_name&"}发言了!<font class=t>("&time&")</font>"
case "查看"
 if Instr(LCase(application("qzld_pingbi")),LCase(";"&aqjh_name&"|"))=0 and Instr(LCase(application("qzld_pingbi")),LCase("|"&aqjh_name&";"))=0 then
  pingbi="##并没有屏蔽任何人也没有被任何人屏蔽呀~！<font class=t>("&time&")</font>"
 else
  pingbi=""
  pingbi1=""
  data1=split(Application("qzld_pingbi"),";")
  data2=UBound(data1)
  for i=1 to data2-1
   data3=split(data1(i),"|")
   if data3(0)=aqjh_name then pingbi=pingbi&"["&data3(1)&"]"
   if data3(1)=aqjh_name then pingbi1=pingbi1&"["&data3(0)&"]"
   erase data3
  next 
  erase data1
  if pingbi<>"" then pingbi="##把："&pingbi&"等屏蔽了！"
  if pingbi1<>"" then pingbi=pingbi&"##被："&pingbi1&"等屏蔽了！"
  pingbi=pingbi&"<font class=t>("&time&")</font>"
 end if
case else
 Response.Write "<script Language=javascript>alert('请不要发不明命令！');</script>"
 response.end
end select
end function
%>

