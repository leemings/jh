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
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以国家之间挑衅！');}</script>"
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
says="<font color=green>【国家挑衅】</font><font color=" & saycolor & ">"+guoopen(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'国家挑战
function mpopen(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 职位,国家,门派 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,3,3
zhiwei1=rs("职位")
guo1=rs("国家")
pai1=rs("门派")
 if pai1="官府" then
       Response.Write "<script language=JavaScript>{alert('您是官府的人，怎么能参与国家之间的纷争，做好你自己的职责！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if guo1="" or guo1="无" then
       Response.Write "<script language=JavaScript>{alert('你现在又没有加入任何国家,怎么参与国战呀快去加入吧！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if zhiwei1<>"君主" then
       Response.Write "<script language=JavaScript>{alert('您不是帝王,没有权利代表本国参与国战！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
rs.open "select 职位,国家,门派 FROM 用户 WHERE 姓名='" & to1 &"'",conn
zhiwei2=rs("职位")
guo2=rs("国家")
pai2=rs("门派")
 if pai2="官府" then
       Response.Write "<script language=JavaScript>{alert('对方是官府的人,你找死阿,滚远点去！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if guo2="" or guo1="无" then
       Response.Write "<script language=JavaScript>{alert('对方无国无界,你怎么忍心杀这么一个人呢！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if zhiwei2<>"君主" then
       Response.Write "<script language=JavaScript>{alert('对方不是国王,没有权利代表对方国家参与国战！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT d,bh FROM guo WHERE  g='" & guo1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('你好像不是本国的人，捣什么乱那！');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  enemy1=rs("d")
  protect1=rs("bh")
    if InStr("|" & enemy1 & "|", "|" & guo2& "|")>0 then
Response.Write "<script language=JavaScript>{alert('你们已经开战了,不需要再挑战了阿.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="SELECT d,bh FROM guo WHERE  g='" & guo2 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('对方好像不是他们国家中人，还是放过吧');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
 enemy2=rs("d")
  protect2=rs("bh")
   if InStr("|" & enemy2 & "|", "|" & guo1& "|")>0 then
Response.Write "<script language=JavaScript>{alert('对方好像已经和你开战了,不需要再挑战了阿');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       enemys1=enemy1&guo2&"|"
       enemys2=enemy2&guo1&"|"
   conn.execute"update guo set d='"&enemys1&"' where g='" & guo1 & "'"
    conn.execute"update guo set d='"&enemys2&"' where g='" & guo2 & "'"
guoopen=aqjh_name & "</font></b>对<b><font color=black>" & to1 & "</font></b>说： 朕亲率本国百万雄兵来发送英雄帖,盼与你国一拼高下,有胆量的就来，看你们国家有没有这个实力！！！." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>