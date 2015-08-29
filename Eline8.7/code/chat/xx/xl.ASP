<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 金,木,水,火,土,银币 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("金")<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的金属性小于60，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("木")<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的木属性小于60，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("水")<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的水属性小于60，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("火")<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的火属性小于60，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("土")<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的土属性小于60，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
if rs("银币")<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的银币小于60，不能再修炼了！');location.href = 'setwg.asp';</script>"
	Response.End
end if
rs.close
mymp=sjjh_name

rs.open "select a from n where id="& id &" and b='"&mymp&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的轩辕没有这个招试！');history.go(-1);</script>"
	response.end
end if
conn.execute "update n set e=e+1 where  id="&id
conn.execute "update 用户 set 金=金-60 where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 木=木-60 where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 水=水-60 where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 火=火-60 where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 土=土-60 where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 银币=银币-60 where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：修炼轩辕ID"& id &"一次成功！花费金、木、水、火、土及银币各60');location.href = 'setwg.asp';</script>"
%>
