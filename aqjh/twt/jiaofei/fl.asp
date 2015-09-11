<%@ LANGUAGE=VBScript codepage ="936" %>
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
	Response.Write "<script Language=Javascript>alert('提示：你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	response.end
end if
if Application("aqjh_dalie")<>"土匪" then
	Response.Write "<script Language=Javascript>alert('提示：现在还没有土匪？！');window.close();</script>"
	response.end
end if
if aqjh_grade>=7  then
	Response.Write "<script Language=Javascript>alert('提示：官府中人不可剿匪！！');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 内力,体力,银两 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("内力")<1000 or rs("体力")<10000 or rs("银两")<10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：剿匪需要内、体力各10000，白银10000两！！');window.close();</script>"
	response.end
end if
session("dalie")=true
conn.execute "update 用户 set 内力=内力-1000,体力=体力-10000,银两=银两-10000 where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script>
function dalie()
{
location.href='dalie.asp';
}
</script>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="444" height="278">
<param name=movie value="dalie.swf">
<param name=quality value=high>
<embed src="dalie.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
</embed>
</object>