<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select cw from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if isnull(rs("cw")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有宠物，请去购买！');</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物数据出错，请重新购买！');</script>"
	response.end
end if
rs.close
rs.open "select i from b where a='"&zt(1)&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物数据出错，请重新购买！');</script>"
	response.end
end if
mypic=rs("i")
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
if zg>=(4+int(DateDiff("d",zt(2),now())/10)) then
	Response.Write "<script Language=Javascript>alert('提示：您已经照顾"&(4+int(DateDiff("d",zt(2),now())/10))&"次了，请等5小时再操作！');</script>"
	response.end
end if
'名字|类别|生日|体力|攻击|心情|照顾数次|照顾时间
'金龍|金龍|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-30)&"|"&(clng(zt(4))+40)&"|"&(clng(zt(5))-20)&"|"&(zg+1)&"|"&now()
conn.execute "update 用户 set cw='"&temp&"' where  姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：宠物锻炼完成：体:-30 攻:+40 感:-20 !');parent.f3.location.href = 'cw.asp';</script>"
%>
