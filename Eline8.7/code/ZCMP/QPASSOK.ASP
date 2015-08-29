<%@ LANGUAGE=VBScript%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("sjjh_jm")=session("sjjh_jm")+1
if session("sjjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
function chuser(u)
	dim filter,xx,usernameenable,su
	for i=1 to len(u)
		su=mid(u,i,1)
		xx=asc(su)
		zhengchu = -1*xx \ 256
		yushu = -1*xx mod 256
		if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
			chuser=true
			exit function
		end if
	next
	chuser=false
end function
myid=Request.form("id")
wt=lcase(trim(request.form("wt")))
da=lcase(trim(request.form("da")))
for each element in Request.Form
	if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
next
for i=1 to len(wt)
	usernamechr=mid(wt,i,1)
	if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=120"
next
if wt="" or da="" then
	Response.Write "<script Language=Javascript>alert('提示：提示问题或答案有非法字符或为空，请查看！');location.href = 'javascript:history.go(-1)';</script>"
	Response.end
end if
if left(wt,3)="%20" OR InStr(wt,"=")<>0 or InStr(wt,"`")<>0 or InStr(wt,"'")<>0 or InStr(wt," ")<>0 or InStr(wt,"　")<>0 or InStr(wt,"'")<>0 or InStr(wt,chr(34))<>0 or InStr(wt,"\")<>0 or InStr(wt,",")<>0 or InStr(wt,"<")<>0 or InStr(wt,">")<>0 then
	Response.Write "<script Language=Javascript>alert('提示：输入数据有非法字符，请查看！');location.href = 'javascript:history.go(-1)';</script>"
	Response.end
end if
if left(myid,3)="%20" OR InStr(myid,"=")<>0 or InStr(myid,"`")<>0 or InStr(myid,"'")<>0 or InStr(myid," ")<>0 or InStr(myid,"　")<>0 or InStr(myid,"'")<>0 or InStr(myid,chr(34))<>0 or InStr(myid,"\")<>0 or InStr(myid,",")<>0 or InStr(myid,"<")<>0 or InStr(myid,">")<>0 then 
	Response.Write "<script Language=Javascript>alert('提示：输入ID中有非法字符，请查看！');location.href = 'javascript:history.go(-1)';</script>"
	Response.end
end if
if chuser(wt) then 
	Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if instr(wt,"or")<>0 or instr(wt,"'")<>0 or instr(wt,"|")<>0 or instr(wt," ")<>0 then
	Response.Write "<script Language=Javascript>alert('提示：提示问题或答案有非法字符，请查看！');location.href = 'javascript:history.go(-1)';</script>"
	Response.end
end if
if instr(da,"or")<>0 or instr(da,"'")<>0 or instr(da,"|")<>0 or instr(da," ")<>0 then
	Response.Write "<script Language=Javascript>alert('提示：提示问题或答案有非法字符，请查看！');location.href = 'javascript:history.go(-1)';</script>"
	Response.end
end if
if instr(myid,"or")<>0 or instr(myid,"'")<>0 or instr(myid,"|")<>0 or instr(myid," ")<>0 then
	Response.Write "<script Language=Javascript>alert('提示：ID中有非法字符，请查看！');location.href = 'javascript:history.go(-1)';</script>"
	Response.end
end if
if wt="" or da="" or myid="" then
	Response.Write "<script Language=Javascript>alert('提示：ID、提示问题、答案均不能为空，请重新输入！');location.href = 'javascript:history.go(-1)';</script>"
	Response.end
end if
%>
<!--#include file="data.asp"-->
<%
sql="select * from 密码保护 where 用户="&myid
set rs=conn1.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('你尚未申请过密码保护，如果密码已丢失，请找站长联系！');window.close();</script>"
	response.end
end if
a=rs("问题")
b=rs("答案")
rs.close
conn1.close
set conn1=nothing
if wt<>a or da<>b then
	Response.Write "<script Language=Javascript>alert('你的密码提示问题或答案错误，不能找回密码！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'校验用户
rs.open "SELECT * FROM 用户 WHERE ID="&myid,conn
If rs.bof and rs.eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：未找到此注册账号，请找站长联系！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("门派")="官府" or rs("grade")>=6 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你是官府中人，如果要找回密码请找站长！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
name=rs("姓名")
conn.execute "update 用户 set 密码='e10adc3949ba59abbe56e057f20f883e' where id="&myid
'记录
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& name &"','"& name &"','操作','找回密码！')"
rs.close
set rs=nothing
conn.close
set conn=nothing
'记录
%>
<html>
<head>
<title>无标题文档</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#006699" text="#000000">
<p align="center"><b><font color="#FFFFFF" size="4">取回密码成功</font></b></p>
<p align="center"><font color="#FFFF00"><b>你的注册ID为：</b></font> <font color="#FF0000"><b><%=myid%></b></font>  
</p>
<p align="center"><b><font color="#FFFF00">你现在的姓名为：</font><font color="#FF0000"><%=name%></font></b></p>
<p align="center"><b><font color="#FFFF00">你现在的密码为：</font><font color="#FFFFFF">123456</font></b></p>
<p align="center"><font color="#FF0000" size="3"><b>请尽快修改你的密码，以防止密码再次丢失或被盗。</b></font></p>
<p align="center">　</p>
<p align="center"><font color="#FFFFFF" size="2">『E线江湖』</font></p>
<p align="center">　</p>
<p align="center">　</p>
</body>
</html>