<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if sjjh_jhdj<15 then
	Response.Write "<script language=JavaScript>{alert('提示：你不够15级，还是先练练技术吧!');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 体力,状态,道德 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<10000 then
    Response.Write "<script language=JavaScript>{alert('提示：你的体力不够1万，先练练吧!');window.close();}</script>"
	Response.End 
end if
if rs("状态")="死" then
    Response.Write "<script language=JavaScript>{alert('提示：你已经死了，先去复活吧!');window.close();}</script>"
	Response.End 
end if
if rs("道德")<500 then
    Response.Write "<script language=JavaScript>{alert('提示：你的道德不够500，谁敢相信你?');window.close();}</script>"
	Response.End 
end if
Session("diaoyu")=now()
Response.Redirect "hit.asp"
%>