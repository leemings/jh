<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'新人福神
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=Session("aqjh_name")
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
says="<font color=red>【新人冰神】<font color=" & saycolor & ">"+xrfs()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function xrfs()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if towho=Application("aqjh_automanname") or towho="大家" then
 towho=aqjh_name
end if
if aqjh_name=towho then
'对自己
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if Session("xrfs")=true then
    Response.Write "<script language=JavaScript>{alert('提示：您已经有冰神帮助了，不要再来了！');}</script>"
    Response.End
end if
if rs("times")>4 then
    Response.Write "<script language=JavaScript>{alert('提示：只有新人首次登陆才能有冰神帮助！');}</script>"
    Response.End
else
    conn.Execute ("update 用户 set sl='冰神',slsj=now()+1,times=times+3 where 姓名='" & aqjh_name &"'")
    Session("xrfs")=true
end if
xrfs="[##]初次来快乐江湖，突然出现一道金光，神灵赋在身上了！获得三天冰神护身！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
else
'对别人
rs.open "select * FROM 用户 WHERE 姓名='" & towho&"'",conn,2,2
if rs("times")<>3 then
    Response.Write "<script language=JavaScript>{alert('提示：谢谢你，但我已经不需要冰神帮助了！');}</script>"
    Response.End
else
    conn.Execute ("update 用户 set sl='冰神',slsj=now()+1,times=times+3 where 姓名='" & towho &"'")
    conn.Execute ("update 用户 set 道德=道德+100 where 姓名='" & aqjh_name &"'")
end if
xrfs="[%%]初次来快乐江湖，[##]对[%%]使用了新人冰神，突然出现一道金光，神灵赋在[%%]身上了！获得一天快乐之神护身，泡点速度2.5倍！[%%]连声道谢！[##]道德上升100点！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end if
end function
%>