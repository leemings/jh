<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'对赌银两♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不能够赌博！');}</script>"
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
call dianzan(towho)
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
says=duidu(mid(says,i+1),towho)+"</font>"
call chatsay("对赌银两",towhoway,towho,saycolor,addwordcolor,addsays,says)

'对赌银两
'命运江湖首创
function duidu(fn1,to1)
if not isnumeric(fn1) then
	Response.Write "<script language=JavaScript>{alert('["&fn1&"]操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
fn1=int(abs(fn1))
if fn1<10000 or fn1>3000000 then
	Response.Write "<script language=JavaScript>{alert('提示：对赌银两应大于10万，小于300万！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 银两 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2

if rs("银两")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的银两好象不够呀！怎么跟人家赌??');}</script>"
	Response.End
end if
rs.close
rs.open "select 银两 from 用户 where 姓名='"&to1&"'",conn,2,2

if rs("银两")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["&to1&"]你好象没有这么多的银子啊?怎么赌???!');}</script>"
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
'duidu="[##]向{%%}提出对赌邀请，赌注为："&fn1&"银子，哈哈,我们来玩一把 ,<input  type=button value='谁怕谁,来吧!' 'onClick=javascript:duidu"&regjm&".disabled=1;window.open('duidu_add.asp?name="&sjjh_name&"&fn1="&fn1&"&to1="&to1&"&m1="&m1&"&m2="&m2&"&m3="&m3&"','d') name=duidu"&regjm&">"
'duidu="<form method=POST target=d name=duidu"&regjm&" 'action=duidu_add.asp?name="&sjjh_name&"&fn1="&fn1&"&toname="&to1&"&m1="&m1&"&m2="&m2&"&m3="&m3&"><font color=green>【对赌银两】<font color=" & saycolor & ">[##]向{%%}提出对赌邀请，赌注为："&fn1&"银子，哈哈,我们来玩一把!<input type=submit onClick=javascript:duidu"&regjm&".disabled=1;this.document.duidu"&regjm&".submit() name=duidu"&regjm&" value=谁怕谁,来吧!></form>"

duidu="<form method=POST target=d name=duidu"&regjm&" action=duidu_add.asp?name="&urlencoding(sjjh_name)&"&fn1="&fn1&"&toname="&urlencoding(to1)&"&m1="&m1&"&m2="&m2&"&m3="&m3&"><font color=green>【对赌银两】<font color=" & saycolor & ">[##]向{%%}提出对赌请邀请，赌注为："&fn1&"银子，哈哈,我们来玩一把!<input type=submit onClick=javascript:duidu"&regjm&".disabled=1;this.document.duidu"&regjm&".submit() name=duidu"&regjm&" value=谁怕谁,来吧!></form>"
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

