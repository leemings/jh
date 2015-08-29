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
		Response.Write "<script Language=Javascript>alert('提示：是不是想不救自己呀？连大名和进门口令都不报！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
inroom=i
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
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
if rs("银两")<50000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你身上没带那么多钱呀？怎么办？');}</script>"
	Response.End
end if
if rs("法力")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：不好意思,你法力还不足5000点？');}</script>"
	Response.End
end if
if rs("金币")<20 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：没金币还想保释，好好作牢吧！');}</script>"
	Response.End
end if
if rs("状态")<>"眠" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：有没有搞错？你没有睡眠呀？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("状态")="眠" then
		conn.execute("update 用户 set 状态='正常' where 姓名='"&name&"'")
			Response.Write "<script Language=Javascript>alert('OK,你花了5000万银两/20金币和使出了5000点法力终于把自己从睡眠中救醒了!');</script>"
	end if
conn.execute "update 用户 set 银两=银两-50000000,法力=法力-5000,金币=金币-20 where 姓名='"&name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing%>
<head>
<title>解救成功</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699">
<div align="center"><p><font color="#FFFFFF" size="2">你花了5000万银两/20金币和使出了5000点法力终于把自己从睡眠中救醒了。<br>
以后千万要小心噢。。。。。</font></p>
<p><input type="button" value="关闭窗口" onclick="window.close()" name="button">
</p></div>