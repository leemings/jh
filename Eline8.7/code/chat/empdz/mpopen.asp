<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'门派挑战♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用门派挑战！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('提示：门派大战只可以在周五21-22点进行，挑战他派只有掌门才有资格！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
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
says="<font color=green>【门派挑战】</font><font color=" & saycolor & ">"+mpopen(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'门派挑战
function mpopen(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 身份,门派 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,3,3
shenfen1=rs("身份")
men1=rs("门派")
 if men1="官府" then
       Response.Write "<script language=JavaScript>{alert('您是官府的人，怎么能跟江湖门派结怨你疯了吗！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men1="" or men1="无" then
       Response.Write "<script language=JavaScript>{alert('您无门无派，怎么参与门派开战呀快去加入门派吧！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen1<>"掌门" then
       Response.Write "<script language=JavaScript>{alert('您不是掌门，没有权利代表本派参与门派挑战！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
rs.open "select 身份,门派 FROM 用户 WHERE 姓名='" & to1 &"'",conn
shenfen2=rs("身份")
men2=rs("门派")
 if men2="官府" then
       Response.Write "<script language=JavaScript>{alert('对方是官府的人，你找死啊，滚远点去！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men2="" or men1="无" then
       Response.Write "<script language=JavaScript>{alert('对方无门无派，怎么跟你门派挑战呀！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen2<>"掌门" then
       Response.Write "<script language=JavaScript>{alert('对方不是掌门，没有权利代表对方门派参与门派挑战！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT q,f FROM p WHERE  a='" & men1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('你好像不是门派中人');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  shidi1=rs("q")
  baohu1=rs("f")
    if InStr("|" & shidi1 & "|", "|" & men2& "|")>0 then
Response.Write "<script language=JavaScript>{alert('你们已经开战了，不需要再挑战了阿.');}</script>"
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
   if InStr("|" & shidi2 & "|", "|" & men1& "|")>0 then
Response.Write "<script language=JavaScript>{alert('对方好像已经和你开战了，不需要再挑战了阿');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       shidis1=shidi1&men2&"|"
       shidis2=shidi2&men1&"|"
   conn.execute"update p set q='"&shidis1&"' where a='" & men1 & "'"
    conn.execute"update p set q='"&shidis2&"' where a='" & men2 & "'"
mpopen=sjjh_name & "</font></b>对<b><font color=black>" & to1 & "</font></b>说：我亲率门下弟子发送战帖，盼与你门派挑战，完成本门一统江湖大业！" 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing

       
end function 
%>
