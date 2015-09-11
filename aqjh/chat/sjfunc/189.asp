<!--#include file="sjfunc.asp"-->
<%'赌轻功
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不能够赌博！');}</script>"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=duidu2(mid(says,i+1),towho)+"</font>"
call chatsay("赌轻功",towhoway,towho,saycolor,addwordcolor,addsays,says)
'赌轻功
function duidu2(fn1,to1)
if not isnumeric(fn1) then
	Response.Write "<script language=JavaScript>{alert('["&fn1&"]操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
fn1=int(abs(fn1))
if fn1<100 or fn1>10000 then
	Response.Write "<script language=JavaScript>{alert('提示：对赌轻功应大于100，小于10000！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 轻功 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("轻功")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的轻功好像不够呀！');}</script>"
	Response.End
end if
rs.close
rs.open "select 轻功 from 用户 where 姓名='"&to1&"'",conn,2,2
if rs("轻功")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["&to1&"]好像没有这么多的轻功！');}</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
m1 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
regjm=int(rnd*3348998)+1
duidu2="<form method=POST target=d name=duidu2"&regjm&" action=duidu2.asp?name="&urlencoding(aqjh_name)&"&fn1="&fn1&"&toname="&urlencoding(to1)&"&m1="&m1&"&m2="&m2&"&m3="&m3&"><font color=green>【对赌法力】<font color=" & saycolor & ">[##]向{%%}提出对赌请邀请，赌注为："&fn1&"点轻功，<input type=submit onClick=javascript:duidu2"&regjm&".disabled=1;this.document.duidu2"&regjm&".submit() name=duidu2"&regjm&" value=赌轻功></form>"
end function

function urlencoding(vstrin)
    strreturn = ""
    for i = 1 to len(vstrin)
        thischr = mid(vstrin,i,1)
        if abs(asc(thischr)) < &hff then
            strreturn = strreturn & thischr
        else
            innercode = asc(thischr)
            if innercode < 0 then
                innercode = innercode + &h10000
            end if
            hight8 = (innercode  and &hff00)\ &hff
            low8 = innercode and &hff
            strreturn = strreturn & "%" & hex(hight8) &  "%" & hex(low8)
        end if
    next
    urlencoding = strreturn
end function

%>
