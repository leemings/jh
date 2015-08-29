<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if request.form("submit")<>"还款" then
	dk=request.form("dk")
	if dk="" then Response.Redirect "daikuan.asp"
	if not isnumeric(dk) then response.redirect"../error.asp?id=464"
	dk=int(dk)
	if dk<1 then dk=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
If Rs.Bof OR Rs.Eof Then 
	set rs=nothing
	rs.Close
	conn.Close
	set conn=nothing
	Response.Redirect "../error.asp?id=440"
end if
	allvalue=rs("allvalue")
	bigdk=allvalue*100
	if dk-bigdk>0 then 
		set rs=nothing
		rs.Close
		conn.Close
		set conn=nothing
		Response.Redirect "../error.asp?id=482"
	end if
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
rs.close
rs.open "select * from k where a='"&sjjh_name&"'",conn
	'贷款操作项
	if rs.BOF or rs.EOF then
		rs.close
		rs.open "select * from k where d=true",conn
			if rs.BOF or rs.EOF then
				conn.Execute "insert into k(a,b,c,d)values('"&sjjh_name&"',"&sqlstr(sj)&","&dk&",false)"
			else
				id=rs("id")
				conn.execute "update k set b="&sqlstr(sj)&",c="&dk&" ,a='"&sjjh_name&"',d=false where id="&id
			end if
	else
			if rs("d")=true then
				'如果有记录则程式将更新其记录
				conn.execute "update k set b="&sqlstr(sj)&",c="&dk&",d=false where a='"&sjjh_name&"'"
			else
				Response.Redirect "../error.asp?id=483"
			end if
	end if
conn.execute "update 用户 set 银两=银两+"&dk&" where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.Close
set conn=nothing
Response.Redirect "daikuan.asp"
'贷款操作项在此结束
	Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
	End Function

else
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select * from k where a='"&sjjh_name&"'",conn
	dk=rs("c")
	rs.close
	rs.open "select * from 用户 where 姓名='"&sjjh_name&"'"
	qian=rs("银两")
	q=qian-(dk*1.5)
if q<0 then
	response.redirect "daikuan.asp"
	response.end
end if
conn.execute "update k set d=true where a='"&sjjh_name&"'"
conn.execute "update 用户 set 银两=银两-"&dk*1.5&" where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "daikuan.asp"
end if
%>