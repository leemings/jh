<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<!--#include file="../../chk.asp"-->
<!--#include file="../../mywp.asp"-->
<body bgcolor="#000000">
<%
call chkpost()
wpname=LCase(Trim(Request.form("wpname")))
wpsl=LCase(Trim(Request.Form("sl")))
filedname=LCase(Trim(Request.Form("filedname")))
if filedname<>"w1" and filedname<>"w2" and filedname<>"w3" and filedname<>"w4" and filedname<>"w5" and filedname<>"w6" and filedname<>"w7" and filedname<>"w8" then
	Response.Write "<script language=javascript>{alert('数据出错！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
if not isnumeric(wpsl) then
	Response.Write "<script language=javascript>{alert('物品数量请使用半角数字填写！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
wpsl=abs(int(clng(wpsl)))
if wpname="" or instr(wpname,"|")<>0 or instr(wpname,";")<>0 or instr(wpname,">")<>0 or instr(wpname,"<")<>0 or instr(wpname,"=")<>0 or instr(wpname,"(")<>0 or instr(wpname,")")<>0 or instr(wpname,"or")<>0 or instr(wpname,"union")<>0 or instr(wpname,chr(34))<>0 or instr(wpname,chr(39))<>0 or instr(wpname,chr(13))<>0 or instr(wpname,"&")<>0 then
	Response.Write "<script language=javascript>{alert('物品名数据非法！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,会员等级,"&filedname&" from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('你不是江湖中人！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
myid=rs("id")
hydj=rs("会员等级")
mywp=trim(rs(filedname))
rs.close
sjsl=mywpsl(mywp,wpname)
if sjsl<=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('你并没有这样东西！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End
end if
if sjsl<wpsl then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('你的"&wpname&"只有"&sjsl&"个，不够"&wpsl&"个！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End
end if
wpzs=10
wpzs=wpzs+hydj*5	'得出最多可存放几种物品
dim temper
set temper=conn.execute("select count(*) as 数量 from w where a="&myid)
boxwpsl=temper("数量")	'现在总共存放了几种物品
set temper=nothing
rs.open "select * from w where b='"&wpname&"' and a="&myid,conn,2,2
if not (rs.eof or rs.bof) then
	sql="update w set c=c+"&wpsl&" where b='"&wpname&"' and a="&myid
else
	if boxwpsl>=wpzs then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>{alert('你的储物箱已经满了，放不下任何东西了。');location.href = 'javascript:history.go(-1)';}</script>" 
		Response.End
	end if
	sql="insert into w(a,b,c,d) values ("&myid&",'"&wpname&"',"&wpsl&",'"&filedname&"')"
end if
rs.close
newtemp=abate(mywp,wpname,wpsl)
conn.execute "update 用户 set "&filedname&"='"&newtemp&"' where 姓名='"&aqjh_name&"'"
set rs=conn.execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Write "<Script Language=JavaScript>{parent.wupin.document.location.reload();}</Script>"
Response.Redirect "bxg.asp"
%>