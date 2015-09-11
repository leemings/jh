<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以进行门派求和！');}</script>"
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
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="大家" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>【门派求和】</font><font color=" & saycolor & ">"+mpopen1(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'门派求和
function mpopen1(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
sql="select 身份,门派 from 用户 where 姓名='" &aqjh_name&"'"
set rs=conn.execute(sql)
life1=rs("身份")
pai1=rs("门派")
 if pai1="官府" then
       Response.Write "<script language=JavaScript>{alert('您是官府的人，怎么可能与其他门派求和！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if pai1="" or pai1="无" then
       Response.Write "<script language=JavaScript>{alert('您无门无派,怎么参与门派求和呀！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if life1<>"掌门" then
       Response.Write "<script language=JavaScript>{alert('您不是掌门,没有权利代表本派参与门派求和！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

sql="select 身份,门派 from 用户 where 姓名='" &to1&"'"
set rs=conn.execute(sql)
life2=rs("身份")
pai2=rs("门派")
 if pai2="官府" then
       Response.Write "<script language=JavaScript>{alert('对方是官府的人,这么可能有跟你挑战！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if pai2="" or pai1="无" then
       Response.Write "<script language=JavaScript>{alert('对方无门无派,怎么跟你门派求和呀！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if life2<>"掌门" then
       Response.Write "<script language=JavaScript>{alert('对方不是掌门,没有权利代表对方门派参与门派求和！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT q,h,f FROM p WHERE  a='" & pai1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是门派中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  enemy1=rs("q")
  protect1=rs("f")
  kunjin1=rs("h")
  if protect1=1 then
       Response.Write "<script language=JavaScript>{alert('你们正接受官府保护,怎么向人家求和');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

  if kunjin1<=xiamoney then
       Response.Write "<script language=JavaScript>{alert('你们有这么多的库金吗?');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

    if InStr("|" & enemy1 & "|", "|" & pai2& "|")=0 then
Response.Write "<script language=JavaScript>{alert('你们已经求和,不需要再次求和了.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="SELECT q,f FROM p WHERE  a='" & pai2 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是门派中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  enemy2=rs("q")
  protect2=rs("f")
  if protect2=1 then
       Response.Write "<script language=JavaScript>{alert('他们正接受官府保护,你不需要向他们求和');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

   if InStr("|" & enemy2 & "|", "|" & pai1& "|")=0 then
Response.Write "<script language=JavaScript>{alert('你们已经求和,不需要再次求和了.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
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
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
if application("aqjh_mp")<>""  then
dui=split(application("aqjh_mp"),"|")
if DateDiff("s",dui(1),sj2)<60 then
Response.Write "<script Language=Javascript>alert('离前面的求和时间还有" & 60-DateDiff("s",dui(1),sj2) & "秒!');</script>"
response.end
else
Application.Lock
application("aqjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&aqjh_name&"|"&pai1&"|"&pai2
Application.UnLock
mpopen1="["&aqjh_name&"]对<b><font color=black>["&to1&"]</font></b>说：我亲率门下弟子拜送和帖,盼与贵派和睦相处,互惠互利.希望大人有大谅<input type=button value='愿意' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='拒绝' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&pai1&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"">"
end if
else
Application.Lock
application("aqjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&aqjh_name&"|"&pai1&"|"&pai2

Application.UnLock

mpopen1="["&aqjh_name&"]</font></b>对<b><font color=black>["&to1&"]</font></b>说：我亲率门下弟子拜送和帖,盼与贵派和睦相处,互惠互利.希望大人有大谅<input type=button value='愿意' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='拒绝' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&pai1&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" >"
end if
end function

%>