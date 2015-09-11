<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金,木,水,火,土,金币,操作时间 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if DateDiff("s",rs("操作时间"),now())<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('为了防止在乱刷，请等50秒钟操作!');history.go(-1);}</script>"
	Response.End 
end if
if rs("金")<180 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的金属性小于180，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("木")<180 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的木属性小于180，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("水")<180 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的水属性小于180，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("火")<180 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的火属性小于180，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("土")<180 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的土属性小于180，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("金币")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的金币小于10，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
rs.close
mymp=aqjh_name
rs.open "select a from n where id="& id &" and b='"&mymp&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的轩辕没有这个招试！');history.go(-1);</script>"
	response.end
end if
conn.execute "update n set e=e+1 where id="&id
conn.execute "update 用户 set 金=金-180 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 木=木-180 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 水=水-180 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 火=火-180 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 土=土-180 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 金币=金币-10,操作时间=now() where 姓名='" & aqjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：修炼轩辕ID"& id &"一次成功！花费金、木、水、火、土各180个及金币10个');location.href = 'setwg.asp';</script>"
%>