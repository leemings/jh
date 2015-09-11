<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
if Session("aqjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "请进入聊天室再进入当铺操作！！"
window.close()
</script>
<%response.end
end if
wpname=trim(lcase(request("wpname")))
if InStr(wpname,"or")<>0 or InStr(wpname,"=")<>0 or InStr(wpname,"`")<>0 or InStr(wpname,"'")<>0 or InStr(wpname," ")<>0 or InStr(wpname,"　")<>0 or InStr(wpname,"'")<>0 or InStr(wpname,chr(34))<>0 or InStr(wpname,"\")<>0 or InStr(wpname,",")<>0 or InStr(wpname,"<")<>0 or InStr(wpname,">")<>0 then Response.Redirect "../../error.asp?id=54"
lx=lcase(trim(request("lx")))
if InStr(lx,"or")<>0 or InStr(lx,"=")<>0 or InStr(lx,"`")<>0 or InStr(lx,"'")<>0 or InStr(lx," ")<>0 or InStr(lx,"　")<>0 or InStr(lx,"'")<>0 or InStr(lx,chr(34))<>0 or InStr(lx,"\")<>0 or InStr(lx,",")<>0 or InStr(lx,"<")<>0 or InStr(lx,">")<>0 then Response.Redirect "../../error.asp?id=54"
dansl=request("dansl")
allsl=request("allsl")
if not IsNumeric(dansl) or not IsNumeric(allsl) then
	Response.Write "<script language=JavaScript>{alert('警告:参数不正确，请确认填写无误？');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if dansl>allsl then
	Response.Write "<script language=JavaScript>{alert('警告:你有那么多东西可当吗？');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if allsl=0 or dansl=0 then
	Response.Write "<script language=JavaScript>{alert('警告:没事找事吗？你不当来这里干嘛，滚！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,w1,w2,w4,w5,w6,w7 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"秒钟]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
select case lx
case "卡片"
	wpsl=dansl
	zstemp=abate(rs("w5"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('没有你想当的这种物品，你想做什么？请跟站长联系！');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/4)
			conn.execute "update 用户 set w5='"&zstemp&"',会员金卡=会员金卡+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"张原值会员金卡"&bb&"元拿到当铺当了会员金卡"&yin&"元，大叫：“真值。。。”<font color=#ff00ff>("&time&")</font>"
		end if
	end if
	
case "药品"
	wpsl=dansl
	zstemp=abate(rs("w1"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('没有你想当的这种物品，你想做什么？请跟站长联系！');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update 用户 set w1='"&zstemp&"',银两=银两+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"个原值银两"&bb&"两拿到当铺当了"&yin&"两，很是开心哟。。。<font color=#ff00ff>("&time&")</font>"
		end if
	end if
	
case "毒药"
	wpsl=dansl
	zstemp=abate(rs("w2"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('没有你想当的这种物品，你想做什么？请跟站长联系！');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update 用户 set w2='"&zstemp&"',银两=银两+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"个原值银两"&bb&"两拿到当铺当了"&yin&"两，很是开心哟。。。<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "暗器"
	wpsl=dansl
	zstemp=abate(rs("w4"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('没有你想当的这种物品，你想做什么？请跟站长联系！');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update 用户 set w4='"&zstemp&"',银两=银两+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"个原值银两"&bb&"两拿到当铺当了"&yin&"两，很是开心哟。。。<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "物品"
	wpsl=dansl
	zstemp=abate(rs("w6"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('没有你想当的这种物品，你想做什么？请跟站长联系！');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update 用户 set w6='"&zstemp&"',银两=银两+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"个原值银两"&bb&"两拿到当铺当了"&yin&"两，很是开心哟。。。<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "鲜花"
	wpsl=dansl
	zstemp=abate(rs("w7"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('没有你想当的这种物品，你想做什么？请跟站长联系！');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update 用户 set w7='"&zstemp&"',银两=银两+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"个原值银两"&bb&"两拿到当铺当了"&yin&"两，很是开心哟。。。<font color=#ff00ff>("&time&")</font>"
		end if
	end if

end select

rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【当铺消息】</b></font>"&dan
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
nowinroom=session("nowinroom")
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
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
Response.Write "<script language=JavaScript>{alert('把自己不用的"&lx&":["&wpname&"]"&wpsl&"个原值"&bb&"拿到当铺当了"&yin&"，开心哟。。。！');location.href = 'dan.asp';}</script>"
'Response.Redirect ""
%>