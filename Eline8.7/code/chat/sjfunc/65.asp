<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'赌博♀wWw.happyjh.com♀
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
'对暂离开、点哑穴判断
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
says="<font color=green>【双人赌博】<font color=" & saycolor & ">"+grdb(mid(says,i+1),towho)+"</font></font>"
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function grdb(fn1,toname)
if fn1<>1 and fn1<>2 and fn1<>3 then
	Response.Write "<script language=JavaScript>{alert('输入错误,应该是1-3的数字！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 配偶,道德,魅力,银两 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,2,2
if rs("道德")<300 or rs("魅力")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('道德与魅力不够300，人家不和你赌！');}</script>"
	Response.End
end if
if rs("银两")<100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('没有10万块钱，不能赌博！');}</script>"
	Response.End
end if
rs.close
rs.open "select 银两 FROM 用户 WHERE 姓名='"&toname&"'",conn,2,2
if rs("银两")<100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('他没有10万块钱！');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 银两=银两-100000 where 姓名='"&sjjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_dubos")=toname
Application.unLock
randomize()
regjm=int(rnd*3348998)
grdb="["&sjjh_name&"]向{"&toname&"}提出赌博要求，赌注10万白花花的银子！<input type=button value='石头' onClick="&chr(34)&"javascript:sitou"&regjm&".disabled=1;jianzi"&regjm&".disabled=1;bu"&regjm&".disabled=1;window.open('dubo.asp?name="&sjjh_name&"&db=1&sjjh_bingwen="&fn1&"','d')"&chr(34)&" name=sitou"&regjm&"><input type=button value='剪子' onClick="&chr(34)&"javascript:sitou"&regjm&".disabled=1;jianzi"&regjm&".disabled=1;bu"&regjm&".disabled=1;window.open('dubo.asp?name="&sjjh_name&"&db=2&sjjh_bingwen="&fn1&"','d')"&chr(34)&" name=jianzi"&regjm&"><input type=button value='布' onClick="&chr(34)&"javascript:sitou"&regjm&".disabled=1;jianzi"&regjm&".disabled=1;bu"&regjm&".disabled=1;window.open('dubo.asp?name="&sjjh_name&"&db=3&sjjh_bingwen="&fn1&"','d')"&chr(34)&" name=bu"&regjm&">"
end function
%>