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
%>
<%
wpname=trim(lcase(request("wpname")))
if InStr(wpname,"or")<>0 or InStr(wpname,"=")<>0 or InStr(wpname,"`")<>0 or InStr(wpname,"'")<>0 or InStr(wpname," ")<>0 or InStr(wpname,"　")<>0 or InStr(wpname,"'")<>0 or InStr(wpname,chr(34))<>0 or InStr(wpname,"\")<>0 or InStr(wpname,",")<>0 or InStr(wpname,"<")<>0 or InStr(wpname,">")<>0 then Response.Redirect "../../error.asp?id=54"
lx=lcase(trim(request("lx")))
if InStr(lx,"or")<>0 or InStr(lx,"=")<>0 or InStr(lx,"`")<>0 or InStr(lx,"'")<>0 or InStr(lx," ")<>0 or InStr(lx,"　")<>0 or InStr(lx,"'")<>0 or InStr(lx,chr(34))<>0 or InStr(lx,"\")<>0 or InStr(lx,",")<>0 or InStr(lx,"<")<>0 or InStr(lx,">")<>0 then Response.Redirect "../../error.asp?id=54"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,w7 from 用户 where 姓名='"&aqjh_name&"'",conn
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
case "冰雪"
	wpsl=mywpsl(rs("w7"),wpname)
	zstemp=del(rs("w7"),wpname)
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
			yin=int(bb/2)
			conn.execute "update 用户 set w7='"&zstemp&"',金币=金币+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"朵原值金币"&bb&"元拿到爱情花店由梦幻网络收购，赚了金币"&yin&"个，大叫：“真值。。。”<font color=#ff00ff>("&time&")</font>"
		end if
	end if
	
case "逸韵"
	wpsl=mywpsl(rs("w7"),wpname)
	zstemp=del(rs("w7"),wpname)
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
			yin=int(bb/2)
			conn.execute "update 用户 set w7='"&zstemp&"',金币=金币+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"朵原值金币"&bb&"元拿到爱情花店由梦幻网络收购，赚了金币"&yin&"个，大叫：“真值。。。”<font color=#ff00ff>("&time&")</font>"
		end if
	end if
	
case "香芸"
	wpsl=mywpsl(rs("w7"),wpname)
	zstemp=del(rs("w7"),wpname)
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
			yin=int(bb/2)
			conn.execute "update 用户 set w7='"&zstemp&"',金币=金币+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"朵原值金币"&bb&"元拿到爱情花店由梦幻网络收购，赚了金币"&yin&"个，大叫：“真值。。。”<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "忧梦"
	wpsl=mywpsl(rs("w7"),wpname)
	zstemp=del(rs("w7"),wpname)
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
			yin=int(bb/2)
			conn.execute "update 用户 set w7='"&zstemp&"',金币=金币+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"朵原值金币"&bb&"元拿到爱情花店由梦幻网络收购，赚了金币"&yin&"个，大叫：“真值。。。”<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "烈焰"
	wpsl=mywpsl(rs("w7"),wpname)
	zstemp=del(rs("w7"),wpname)
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
			yin=int(bb/2)
			conn.execute "update 用户 set w7='"&zstemp&"',金币=金币+" & yin & ",操作时间=now() where 姓名='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}把自己不用的"&lx&":["&wpname&"]"&wpsl&"朵原值银两"&bb&"元拿到爱情花店由梦幻网络收购，赚了银两"&yin&"元，大叫：“真值。。。”<font color=#ff00ff>("&time&")</font>"
		end if
	end if
	case "牡丹","雏菊","郁金香","此情不渝","醉百合","昨日香","玫瑰","月季","迎风笑","飞杨","满天星","黑郁金香","雪莲花","牵手","蝴蝶兰","火烈鸟","百年好合","马蹄莲","高风亮节","浓情似火","美满幸福","幸福泉源","情意浓浓","健康永远","节日献礼","梦中情人","温馨的爱","灿烂人生","此情可待","真情歉意","幸福安康","粉色舞蹈","水晶之恋","爱的诉说","无言感激","红袖留香","我的牵挂","风情玫瑰","快乐时光","情定一生","一往情深","两情相悦","祝你快乐","特别的爱","青春永存","平安快乐","生意兴隆","爱到永远","千里相思"
	wpsl=mywpsl(rs("w7"),wpname)
	zstemp=del(rs("w7"),wpname)
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

says="<font color=red><b>【梦幻网络收花】</b></font>"&dan
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
Response.Write "<script language=JavaScript>{alert('把自己不用的"&lx&":["&wpname&"]"&wpsl&"个原值"&bb&"拿到花店卖了"&yin&"，开心哟。。。！');location.href = 'shouhua.asp';}</script>"
'Response.Redirect ""
%>