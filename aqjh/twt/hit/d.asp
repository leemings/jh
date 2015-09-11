<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if aqjh_jhdj<15 then
	Response.Write "<script language=JavaScript>{alert('提示：你不够15级，还是先练练技术吧!');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 体力,状态,道德 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
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
Response.Redirect "diao.asp"
%>