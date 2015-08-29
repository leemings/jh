<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"

if Instr(LCase(Application("sjjh_useronlinename"&nowinroom))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

id=request("id")
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

rs.open "SELECT ID,银两 FROM 用户 WHERE 姓名='" & sjjh_name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('江湖中找不到你的资料！');window.close();</script>"
	response.end
end if
sjjh_id=rs("id")
yinliang=rs("银两")
rs.close
%><!--#include file="jiu.asp"--><%
sql="select * from 名妓 where ID="& id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('你弄错了吧，本院没有这个姑娘呀！');window.close();</script>"
	response.end
end if
meimao=rs("美貌度")
mingji=rs("姓名")
yin=meimao*32000
if yinliang<yin then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('替小姐"& mingji &"赎身的价格是"& yin &"两银子，你没有这么多钱！');window.close();</script>"
	response.end
end if
conn.execute "update 用户 set 银两=银两-" & yin & ",道德=道德+300 where 姓名='"& sjjh_name &"'"
connt.execute("delete * from 名妓 where ID="&id)
rs.close
set rs=nothing
conn.close
set conn=nothing
connt.close
set connt=nothing
Response.Write "<script Language=Javascript>alert('这是小姐"& mingjh &"的卖身契，她现在自由了，你真是个好人，好人会有好报的！');window.close();</script>"
response.end
%>
