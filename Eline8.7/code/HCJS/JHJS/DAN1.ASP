<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER")
if InStr(http,"hcjs/jhjs")=1 then
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>"
Response.End
end if
danlb=lcase(trim(request("lb")))
danname=lcase(trim(request("name")))
dansl=clng((trim(request("sl"))))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select "&danlb&" from 用户 where 姓名='"&sjjh_name&"'",conn
'Response.Write danlb&"<br>"&danname&"<br>"&dansl&"abc"&rs(danlb)
'Response.End
dantemp=abate(rs(danlb),danname,dansl)
rs.close
rs.open "select * FROM b WHERE a='" & danname &"'",conn
If Rs.Bof OR Rs.Eof then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！！');}</script>"
Response.End
end if
if rs("b")="鲜花" then
miniwj=20000
else
miniwj=(abs(rs("d"))+abs(rs("e"))+abs(rs("f"))+abs(rs("g")))*5
end if
if rs("b")="鲜花" then
miniwj=20000*dansl
elseif rs("b")="花园" then
miniwj=200000*dansl
else
miniwj=(abs(rs("d"))+abs(rs("e"))+abs(rs("f"))+abs(rs("g")))*5*dansl
end if
conn.execute "update 用户 set "& danlb &"='"&dantemp&"',银两=银两+"&miniwj&" where  姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
nowinroom=session("nowinroom")
dan="把物品：<font color=blue>"&danname&"</font>在当铺当出<font color=blue>"& dansl &"</font>,获得银两:<font color=blue>"& miniwj &"</font>两……"
says="<font color=#ff0000>【消息】" & sjjh_name & ""& dan &"</font>"			'聊天数据
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Response.Write "<script language=javascript>{alert('当出成功，获得银两:"&miniwj&"两!');parent.history.go(-1);}</script>"
Response.End
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
%>