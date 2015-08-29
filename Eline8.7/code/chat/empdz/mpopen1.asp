<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'门派求和♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以进行门派求和！');}</script>"
	Response.End
end if
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
'对暂离开、点哑穴判断
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="大家" or towho=application("sjjh_automanname") or towho=sjjh_name then
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
connstr=Application("sjjh_usermdb")
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('提示：门派大战只可以在周五21-22点进行！并且只有堂主以上才有资格，请提前和掌门准备好！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
sql="select 身份,门派 from 用户 where 姓名='" &sjjh_name&"'"
set rs=conn.execute(sql)
shenfen1=rs("身份")
men1=rs("门派")
 if men1="官府" then
       Response.Write "<script language=JavaScript>{alert('您是官府的人，怎么可能与其他门派求和！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men1="" or men1="无" then
       Response.Write "<script language=JavaScript>{alert('您无门无派，怎么参与门派求和呀！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen1<>"掌门" then
       Response.Write "<script language=JavaScript>{alert('您不是掌门，没有权利代表本派参与门派求和！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

sql="select 身份,门派 from 用户 where 姓名='" &to1&"'"
set rs=conn.execute(sql)
shenfen2=rs("身份")
men2=rs("门派")
 if men2="官府" then
       Response.Write "<script language=JavaScript>{alert('对方是官府的人，这么可能有跟你挑战！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men2="" or men1="无" then
       Response.Write "<script language=JavaScript>{alert('对方无门无派,怎么跟你门派求和呀！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen2<>"掌门" then
       Response.Write "<script language=JavaScript>{alert('对方不是掌门，没有权利代表对方门派参与门派求和！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT q,h,f FROM p WHERE  a='" & men1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是门派中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  shidi1=rs("q")
  baohu1=rs("f")
  kunjin1=rs("h")
  if baohu1=1 then
       Response.Write "<script language=JavaScript>{alert('你们正接受官府保护，怎么向人家求和');}</script>"
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

    if InStr("|" & shidi1 & "|", "|" & men2& "|")=0 then
Response.Write "<script language=JavaScript>{alert('你们已经求和，不需要再次求和了.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="SELECT q,f FROM p WHERE  a='" & men2 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是门派中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  shidi2=rs("q")
  baohu2=rs("f")
  if baohu2=1 then
       Response.Write "<script language=JavaScript>{alert('他们正接受官府保护，你不需要向他们求和');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

   if InStr("|" & shidi2 & "|", "|" & men1& "|")=0 then
Response.Write "<script language=JavaScript>{alert('你们已经求和，不需要再次求和了.');}</script>"
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
if application("sjjh_mp")<>""  then
dui=split(application("sjjh_mp"),"|")
if DateDiff("s",dui(1),sj2)<60 then
Response.Write "<script Language=Javascript>alert('离前面的求和时间还有" & 60-DateDiff("s",dui(1),sj2) & "秒!');</script>"
response.end
else
Application.Lock
application("sjjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&sjjh_name&"|"&men1&"|"&men2
Application.UnLock
mpopen1="["&sjjh_name&"]对<b><font color=black>["&to1&"]</font></b>说：我亲率门下弟子拜送和帖，盼与贵派和睦相处，互惠互利，希望大人有大谅<input type=button value='愿意' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='拒绝' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&men1&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"">"
end if
else
Application.Lock
application("sjjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&sjjh_name&"|"&men1&"|"&men2

Application.UnLock

mpopen1="["&sjjh_name&"]</font></b>对<b><font color=black>["&to1&"]</font></b>说：我亲率门下弟子拜送和帖，盼与贵派和睦相处，互惠互利，希望大人有大谅<input type=button value='愿意' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='拒绝' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&men1&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" >"
end if
end function

%>