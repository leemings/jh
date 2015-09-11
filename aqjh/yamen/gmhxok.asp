<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
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
name=lcase(trim(request.form("name")))
pass=request.form("pass")
name1=lcase(trim(request.form("name1")))
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
for i=1 to len(name1)
usernamechr=mid(name1,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=120"
next
if name1="无" or name="无" or name="未定" or name1="未定" then Response.Redirect "../error.asp?id=130"
if left(name1,3)="%20" OR InStr(name1,"=")<>0 or InStr(name1,"`")<>0 or InStr(name1,"'")<>0 or InStr(name1," ")<>0 or InStr(name1,"　")<>0 or InStr(name1,"'")<>0 or InStr(name1,chr(34))<>0 or InStr(name1,"\")<>0 or InStr(name1,",")<>0 or InStr(name1,"<")<>0 or InStr(name1,">")<>0 then Response.Redirect "../error.asp?id=120"
if chuser(name1) then Response.Redirect "../error.asp?id=120"
if instr(name,"or")<>0 or instr(name,"'")<>0 or instr(name,"|")<>0 or instr(name," ")<>0 then Response.Redirect "../error.asp?id=120"
if instr(name1,"or")<>0 or instr(name1,"'")<>0 or instr(name1,"|")<>0 or instr(name1," ")<>0 then Response.Redirect "../error.asp?id=120"
if Instr(name1,"管理")>0 or Instr(name1,"交易人名")>0 or Instr(name1,"朋友名字")>0 or Instr(name1,Application("aqjh_automanname"))>0 or Instr(name1,"阿俊")>0 or Instr(name1,"站长")>0 or Instr(name1,"网管")>0 or Instr(name1,"时代")>0 or Instr(name1,"严")>0 or Instr(name1,"妈")>0 or Instr(name1,"爸")>0 or Instr(name1,"大家")>0 or Instr(name1,"操")>0 then Response.Redirect "../error.asp?id=130"
if name="" or name1="" or pass="" then Response.Redirect "../error.asp?id=56"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(name)&" ")<>0 then Response.Redirect "../error.asp?id=61"
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'校验用户
rs.open "SELECT * FROM 用户 WHERE 姓名='" & name1 & "'",conn
If not rs.bof and not rs.eof  Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=62"
end if
rs.close
rs.open "SELECT * FROM 用户 WHERE id="&myid&" and 姓名='" & name & "' and 密码='" & pass & "'",conn
If Rs.Bof OR Rs.Eof Then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=63"
end if
if rs("银两")<500000000 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=457"
end if
if rs("身份")="掌门" or rs("门派")="官府" or rs("grade")>=6 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=64"
end if
if rs("会员等级")>0 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=65"
end if
sex=rs("性别")
bl=rs("配偶")
conn.Execute "update 用户 set 姓名='" & name1 & "',银两=银两-500000000 where 姓名='" & name & "'"
conn.Execute "update 用户 set 师傅='" & name1 & "' where 师傅='" & name & "'"
if bl<>"无" and bl<>"" then
	conn.Execute "update 用户 set 配偶='" & name1 & "' where 姓名='" & bl & "'"
	if sex="男" then
		conn.execute "update t set b='"& name1 &"' where b='" & name & "'"
	else
		conn.execute "update t set c='"& name1 &"' where c='" & name & "'"
	end if
end if
conn.execute "update k set a='"& name1 &"' where a='" & name & "'"
'记录
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& name &"','"& name1 &"','操作','改换名字！')"
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn1=Server.CreateObject("ADODB.CONNECTION")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../aqjh_data/stock.asp")
conn1.Execute "update 客户 set 帐号='" & name1 & "' where 帐号='" & name & "'"
conn1.Execute "update 股票 set 经营者='" & name1 & "' where 经营者='" & name & "'"
conn1.close
set conn1=nothing
Set conn2=Server.CreateObject("ADODB.CONNECTION")
conn2.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../bbs/data/aqjh_bbs.asp")
conn2.Execute "update Dv_User set username='" & name1 & "' where username='" & name & "'"
conn2.close
set conn2=nothing
'记录
Response.Redirect "../ok.asp?id=111"
%>