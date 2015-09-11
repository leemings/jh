
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
if aqjh_name="" then 
	Response.Redirect "login.asp"
	Response.End
end if

id=Request("id")
act=Request("act")
if (not isnumeric(id))  then 
	Response.Write "<script language=JavaScript>{alert('提示：数据提交有误！');window.close();}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
select case act
case "dq"
	rs.Open "select * from usesp where sid="& id &" and sa='"& aqjh_name &"'", conn, 1,1
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你并有此物品或此物品已经删除！');window.close();}</script>"
		Response.End
	end if
	conn.execute("update usesp set sd=sc,se=false,sa='无' where sid=" & id)
	conn.execute("update use set i=i-1 where id=" & id )
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：ID:"& id &"的物品丢弃成功!');parent.mainFrame.location.href='mysp.asp';}</script>"
	Response.End
case "delphoto"
	rs.Open "select * from use where a='"& aqjh_name &"'", conn, 1,3
	if rs("g")="" or instr(rs("g"),";")=0 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你像册中并没有东西!');}</script>"
		Response.End
	end if
	showdata=Split(rs("g"), ";")
	rs("g")=Replace(rs("g"),showdata(id)&";","")
	if rs("i")>0 then
		rs("i")=rs("i")-1
	end if
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：成功删除一张图片!');parent.mainFrame.location.href='myphoto.asp';}</script>"
	Response.End
case "hzphoto"
	rs.Open "select * from use where a='"& aqjh_name &"'", conn, 1,3
	if rs("g")="" or instr(rs("g"),";")=0 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你像册中并没有东西!');}</script>"
		Response.End
	end if
	showdata=Split(rs("g"), ";")
	rs("f")=showdata(id)
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：成功装备照片!');parent.mainFrame.setcookie('"& showdata(id) &"');parent.mainFrame.location.href='myphoto.asp';}</script>"
	Response.End
end select
%>