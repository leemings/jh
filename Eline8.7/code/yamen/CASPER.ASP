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
		Response.Write "<script Language=Javascript>alert('提示：是不是不想报仇血恨了？连大名和进门口令都不报！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'以上是对联盟网吧判断
wbname=""
wbpf=0
ip=Request.ServerVariables("REMOTE_ADDR")
rs.open "SELECT barname FROM bar WHERE ip='"&ip&"'",conn
if Not(rs.Eof and rs.Bof) then
wbname=rs("barname")
wbpf=1
end if
rs.close
pass=md5(pass)
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
if rs("体力")>-100 and rs("状态")<>"死" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：有没有搞错？这个人还没死啊？？');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("状态")<>"死" and rs("状态")<>"正常" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：此人被系统监禁中……');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("门派")<>"天网" then
	str=",保护=True"
else
	str=",保护=False"
end if
'以下是对杀手的判断
	'网吧联盟判断
if wbpf<>0 then
	conn.execute("update 用户 set 状态='正常',体力=1000,保护=true where 姓名='"&name&"'")
	Response.Write "<script Language=Javascript>alert('你在加盟网吧["&wbname&"上网，复活状态不变！]');</script>"
else
	'会员等级判断
	if rs("会员等级")>1 then
		conn.execute("update 用户 set 状态='正常',体力=1000,保护=true where 姓名='"&name&"'")
		Response.Write "<script Language=Javascript>alert('OK,你是2级会员，所以你复活了什么也没有丢，但是，我们还是不希望你再来了!');</script>"
	else
		conn.execute("update 用户 set 状态='正常',体力=1000,内力=100,银两=100,保护=true where 姓名='"&name&"'")
			Response.Write "<script Language=Javascript>alert('OK,你现在已经复活了，不要再来了啊,如果是2级会员或在加盟网吧上网将不会有任何损失!');</script>"
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
<head>
<title>复活成功</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699">
<div align="center">
  <p><font color="#FFFFFF" size="2">恭喜大侠你成功的复生了，物品[</font><font size="2"><%=wpname%><font color="#FFFFFF">]丢失了，站长帮助你保护了！记住，我让你复生<br>
    你一定要报仇一定要。。。。。</font></font></p>
<p><input type="button" value="关闭窗口" onclick="window.close()" name="button">
</p></div>