<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<!--#include file="../../chk.asp"-->
<!--#include file="../../mywp.asp"-->
<%
response.write "<body bgcolor=#000000>"
call chkpost()
wpid=Request("id")
wpsl=LCase(Trim(Request.Form("sl")))
if not isnumeric(wpsl) then
	Response.Write "<script language=javascript>{alert('物品数量请使用半角数字填写！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
if not isnumeric(wpid) then
	Response.Write "<script language=javascript>{alert('非法数据！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
wpsl=abs(int(clng(wpsl)))
if wpsl=0 then
	Response.Write "<script language=javascript>{alert('请填写要取的物品数量！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from w where id="&wpid,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('你的储物箱中没有这样东西！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
yhid=rs("a")
wpname=rs("b")
wpzs=rs("c")
filedname=rs("d")
if wpsl>wpzs then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('储物箱中的["&wpname&"]只有["&wpzs&"]个，不够"&wpsl&"。');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
rs.close
rs.open "select id,"&filedname&" from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('你不是江湖中人！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
myid=rs("id")
if yhid<>myid then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('这不是你的东西。');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
mywp=trim(rs(filedname))
if mywp<>"" and not isnull(mywp) then
	data1=split(mywp,";")
	data2=UBound(data1)
	if data2>=25 then
		erase data1
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>{alert('你身上所带的此类物品已超过25种，再也放不下了。');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End
	end if
	erase data1
end if
if wpsl=wpzs then
	conn.execute "delete * from w where id="&wpid&" and a="&myid
else
	conn.execute "update w set c=c-"&wpsl&" where id="&wpid&" and a="&myid
end if
temper=add(mywp,wpname,wpsl)
conn.execute "update 用户 set "&filedname&"='"&temper&"' where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<Script Language=JavaScript>{parent.wupin.document.location.reload();}</Script>"
Response.Redirect "bxg.asp"
%>