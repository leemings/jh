<%@ LANGUAGE=VBScript codepage ="936" %>
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
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
vhid=trim(Request.Form("id"))
vhname=trim(Request.Form("vhname"))	 'g
jrtg=trim(Request.Form("jrtg"))		 'd or f
jr=Request.Form("jr")					 'c or e
jrtc=Request.Form("jrtc")
vhname=replace(vhname," ","")
jrtg=replace(jrtg," ","")

if not IsNumeric(vhid) or InStr(vhname,"or")<>0 or InStr(vhname,"'")<>0 or InStr(vhname,"`")<>0 or InStr(vhname,"=")<>0 or InStr(vhname,"-")<>0 or InStr(vhname,",")<>0 or InStr(vhname,"<")<>0 or InStr(vhname,">")<>0 or InStr(vhname,"[")<>0 or InStr(vhname,"]")<>0 or InStr(jrtg,"or")<>0 or InStr(jrtg,"'")<>0 or InStr(jrtg,"`")<>0 or InStr(jrtg,"=")<>0 or InStr(jrtg,"-")<>0 or InStr(jrtg,",")<>0 or InStr(jrtg,"<")<>0 or InStr(jrtg,">")<>0 or InStr(jrtg,"[")<>0  or InStr(jrtg,"]")<>0  or InStr(jr,"or")<>0 or InStr(jr,"'")<>0 or InStr(jr,"`")<>0 or InStr(jr,"=")<>0 or InStr(jr,"-")<>0 or InStr(jr,",")<>0 then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
if InStr(jrtg,"##")=0 or InStr(jrtg,"$$")=0 then
	Response.Write "<script Language=Javascript>alert('提示：进入公告中必须包含“##”和“$$”！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if


Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT b FROM vh where id="&vhid,conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你想做什么？想捣乱吗？！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("b")<>sjjh_name then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？这个又不是你的，想捣乱吗？！');}</script>"
	response.end
end if
rs.close
rs.open "select 操作时间 from 用户 where 姓名='"&sjjh_name&"'",conn
sj=DateDiff("s",rs("操作时间"),now())
if sj<5 then
	s=5-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"秒钟]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
conn.execute "update 用户 set 操作时间=now() where 姓名='"&sjjh_name&"'"
if jrtc then
	conn.execute "update vh set c=false where b='"&sjjh_name&"'"
	conn.execute "update vh set c="&jr&",d='"&jrtg&"',g='"&vhname&"' where id="&vhid
else
	conn.execute "update vh set e=false where b='"&sjjh_name&"'"
	conn.execute "update vh set e="&jr&",f='"&jrtg&"',g='"&vhname&"' where id="&vhid
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
if jrtc then
Response.Write "<script Language=Javascript>alert('提示：更新进入江湖聊天室公告成功！');location.href = 'javascript:history.go(-1)';</script>"
else
Response.Write "<script Language=Javascript>alert('提示：更新退出江湖聊天室公告成功！');location.href = 'javascript:history.go(-1)';</script>"
end if
%>
