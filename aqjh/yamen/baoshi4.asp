<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
myid=Request.form("id")
name=request.form("name")
pass=trim(request.form("pass"))
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
if name="" or pass="" then
		Response.Write "<script Language=Javascript>alert('提示：是不是想一辈子做牢房呀？连大名和进门口令都不报！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
inroom=i
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where id="& myid &" and 姓名='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：有没有搞错？哪有这个人啊？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(pass)<>rs("密码") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：输入密码不对？？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("存款")<30000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你身上没带那么多钱呀？怎么办？');}</script>"
	Response.End
end if
if rs("状态")<>"狱" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：有没有搞错？大牢里面没有你这个人？？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
conn.execute("update 用户 set 状态='正常' where 姓名='"&name&"'")
conn.execute "update 用户 set 存款=存款-30000000 where 姓名='"&name&"'"
Response.Write "<script Language=Javascript>alert('OK,你拿出了3000万银两终于自我保释了,以后少干点坏事吧!');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing%>