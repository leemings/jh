<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');}</script>" 
	Response.End 
end if 
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=LCase(trim(request.querystring("name")))
to1=LCase(trim(request.querystring("to1")))
fn1=abs(int(request.querystring("fn1")))
m1=abs(int(request.querystring("m1")))
m2=abs(int(request.querystring("m2")))
m3=abs(int(request.querystring("m3")))
for each element in request.querystring
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');</script>"
		Response.End
end if
next
if to1<>sjjh_name then
		Response.Write "<script Language=Javascript>alert('提示：["&to1&"]人家也不是要与你赌木属性你筹什么趣！');</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 木 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
if rs("木")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的木好象不够呀！');}</script>"
	Response.End
end if
rs.close
rs.open "select 木 from 用户 where 姓名='"&name&"'",conn,2,2
if rs("木")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["&name&"]好象没有这么多的木属性吧！');}</script>"
	Response.End
end if
randomize()
m11 = Int(6 * Rnd + 1)
m22 = Int(6 * Rnd + 1)
m33 = Int(6 * Rnd + 1)
myds=m11+m22+m33
tods=m1+m2+m3
if myds<>tods then
	if myds>tods then
		temp="<font color=green>【赌木属性】</font>[<font color=red>"&sjjh_name&"</font>]摇出:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"点,{<font color=red>"&name&"</font>}摇出:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"点，<font color=red><b>赢了</b></font>"&fn1&"点!不好意思，谢谢兄弟的厚爱！"
		conn.execute "update 用户 set 木=木+" & (fn1-3) & " where 姓名='"&sjjh_name&"'"		
		conn.execute "update 用户 set 木=木-" & fn1 & " where 姓名='"&name&"'"		
	else
		temp="<font color=green>【赌木属性】</font>[<font color=red>"&sjjh_name&"</font>]摇出:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"点,{<font color=red>"&name&"</font>}摇出:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"点，<font color=red><b>输了</b></font>"&fn1&"点!没关系，下次一定赢你！"
		conn.execute "update 用户 set 木=木-" & fn1 & " where 姓名='"&sjjh_name&"'"
		conn.execute "update 用户 set 木=木+" & (fn1-3) & " where 姓名='"&name&"'"
	end if
else
temp="<font color=green>【赌木属性】</font>[<font color=red>"&sjjh_name&"</font>]摇出:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"点,{<font color=red>"&name&"</font>}摇出:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"点，<font color=red><b>平手</b></font>!伟大的思想总是不谋而合……"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【江湖消息】</b></font>"&temp
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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