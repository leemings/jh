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
conn.open Application("aqjh_usermdb")
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
	'会员等级判断
	if rs("会员等级")>1 then
		conn.execute("update 用户 set 状态='正常',体力=1000,保护=true where 姓名='"&name&"'")
		Response.Write "<script Language=Javascript>alert('OK,你是2级会员，所以你复活了什么也没有丢，但是，我们还是不希望你再来了!');</script>"
	else
		conn.execute("update 用户 set 状态='正常',体力=1000,内力=100,银两=100,保护=true where 姓名='"&name&"'")
			Response.Write "<script Language=Javascript>alert('OK,你现在已经复活了，不要再来了啊,如果是2级以上会员不会有任何损失!');</script>"
	end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('恭喜大侠你成功的复生了!记住，你一定要报仇啊!');window.close();</script>"
says="<font color=red><b>【浴火重生】</b>["&name&"]</font><font color=blue>路遇仙人施法，终于再次获得新生。“杀我的人听好了，我一定要找你们复仇！！！</font>"
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & "系统机器" & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>