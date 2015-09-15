<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
aqjh_name=aqjh_name
who=Trim(Request.Form("who"))
if who="" then who=aqjh_name
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
rs.open "SELECT * FROM 用户 where 姓名='" & aqjh_name &"'",conn
if rs.bof or rs.eof then
	response.write "你不是江湖中人"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("grade")>5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你们夫妻有一个是管理员，禁止使用！');history.go(-1);</script>"
	response.end
end if
peiou=rs("配偶")
rs.close
rs.open "SELECT * FROM 用户 where 姓名='" & peiou & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('没有配偶呀！！');history.go(-1);</script>"
	response.end
end if
if rs("grade")>5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你们夫妻有一个是管理员，禁止使用！');history.go(-1);</script>"
	response.end
end if
rs.close
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>合体技</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center">
<p><br>
<font color="#FFFFFF"><span style='font-size:9pt'>〖合体技〗</span></font></font><br>
</font> <font color="#FFFFFF"><br>
攻击力=男方杀伤力+女方杀伤力</font></p>
<form method=POST action='compact1.asp'>
<font color="#FFFFFF"></font><font color="#FF0033">攻击对象</font><br>
<select name="to1">
<%
for i=0 to x
if show(i)<>peiou and  show(i)<>aqjh_name Then
%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%
end if
next
%>
</select>
<font color="#FFFFFF"><br>
<br>
<font color="#FF0033">合体技一览</font><br>
<%
rs.open "SELECT * FROM t where b='" & aqjh_name & "' or c='" & aqjh_name & "'"
if rs.eof and rs.bof then
	response.write "<p align='center'>你没有任何合体技！<br>请自行到合体特技处创建。</p>"
else%>
	<select class="smallSel" name="at"><%
	do while not rs.eof
	sel="selected"
	response.write "<option " & sel & " value='"+CStr(rs("a"))+"' >"+rs("a")+""+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	%>
	</select>
	<br>
	<br>
	</font><font color=#ffffff>
	<input type=submit value="发 招" name=submit>
	</font>
<%end if%>
<br></form>	</div></body></html>
