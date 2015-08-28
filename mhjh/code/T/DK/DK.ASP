<%
Response.Buffer=true
username=session("yx8_mhjh_username")
grade=session("yx8_mhjh_usergrade")
if username="" then Response.Redirect "../../error.asp?id=016"
if grade<20 then 
response.write "江湖小辈，本钱庄可不借！到时老本都没了，找谁？"
response.write "<br><a href=daikuan.asp>返回</a>"
response.end
end if
if request.form("submit")<>"还款" then
dk=request.form("dk")
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"&")<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"") then Response.Redirect "../../error.asp?id=056"
next
if dk="" then Response.Redirect "daikuan.asp"
if not isnumeric(dk) then response.redirect"../../error.asp?id=517"
dk=int(dk)
if dk<1 then dk=1
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 积分,银两 FROM 用户 WHERE 姓名='"&username&"'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then Response.Redirect "../../error.asp?id=016"
allvalue=rs("积分")
yinliang=rs("银两")
rs.Close
conn.Close 
set rs=nothing
set conn=nothing
bigdk=int(yinliang*0.05-allvalue*4)
if bigdk<0 then bigdk=int(yinliang*0.05)
if dk-bigdk>0 then Response.Redirect "../../error.asp?id=030"
	n=Year(date())
	y=Month(date())
	r=Day(date())
	s=Hour(time())
	f=Minute(time())
	m=Second(time())
	if len(y)=1 then y="0" & y
	if len(r)=1 then r="0" & r
	if len(s)=1 then s="0" & s
	if len(f)=1 then f="0" & f
	if len(m)=1 then m="0" & m
	sj=n & "-" & y & "-" & r 
%>
<!--#include file="data.asp"--><%
sql="select * from 贷款 where 贷款人='"&username&"'"
Set rst=connt.execute(sql)
'贷款操作项
if rst.BOF or rst.EOF then
	sql="select * from 贷款 where 还贷记录=true"
	set rst=connt.Execute(sql)
		if rst.BOF or rst.EOF then
		sql="insert into 贷款(贷款人,贷款日期,贷款金额,还贷记录)values('"&username&"',"&sqlstr(sj)&","&dk&",false)"
		set rst=connt.Execute(sql)
		else 
		id=rst("id")
		sql="update 贷款 set 贷款日期="&sqlstr(sj)&",贷款金额="&dk&" ,贷款人='"&username&"',还贷记录=false where id="&id
		set rst=connt.execute(sql)
		end if
else
	if rst("还贷记录")=true then
	'如果有记录则程式将更新其记录
	sql="update 贷款 set 贷款日期="&sqlstr(sj)&",贷款金额="&dk&",还贷记录=false where 贷款人='"&username&"'"
	set rst=connt.execute(sql)
	else
	Response.Redirect "../../error.asp?id=519"
	end if
	
end if
connt.Close
set connt=nothing

			sql="update 用户 set 银两=银两+"&dk&" where 姓名='"&username&"'"
			set conn=server.CreateObject("adodb.connection")
			conn.Mode=16
			conn.IsolationLevel=256
			conn.Open Application("yx8_mhjh_connstr")
			conn.Execute(sql)
			conn.Close
			set conn=nothing
		response.write "<script language='javascript'>alert ('恭喜，你的贷款申请被接纳了！！！');location.href='daikuan.asp';</script>"

'贷款操作项在此结束
Function SqlStr(data)
 SqlStr="'" & Replace(data,"'","''") & "'"
End Function

else
%>
<!--#include file="data.asp"--><%
sql="select * from 贷款 where 贷款人='"&username&"'"
set rst=connt.execute(sql)
dk=rst("贷款金额")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 银两,会员 from 用户 where 姓名='"&username&"'"
set rs=conn.execute(sql)
qian=rs("银两")
hy=rs("会员")
rs.Close
if hy=true then
   q=qian-(dk*1.2)
else
   q=qian-(dk*1.5)
end if
if q<0 then
response.redirect "daikuan.asp"
response.end
end if
sql="update 贷款 set 还贷记录=true where 贷款人='"&username&"'"
set rst=connt.execute(sql)
connt.close
set connt=nothing
if hy=true then
sql="update 用户 set 银两=银两-"&dk*1.2&" where 姓名='"&username&"'"
else
sql="update 用户 set 银两=银两-"&dk*1.5&" where 姓名='"&username&"'"
end if
set rs=conn.execute(sql)
conn.close
set conn=nothing
Response.write "<script language='javascript'>alert ('你的贷款已还清，欢迎再次光临！！！');location.href='daikuan.asp';</script>"

end if
%>
