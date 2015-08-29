<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=LCase(trim(Request.QueryString("name")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w1,boy,boysex,配偶 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
w1=rs("w1")
if iswp(w1,"白饭")=0 then
	Response.Write "<script Language=Javascript>alert('提示：您没白饭，怎么喂养呀！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
zz=rs("配偶")

myhua=rs("w1")
if isnull(rs("boy")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有小孩，赶快生育一个吧！');</script>"
	response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩数据出错，请重新购买！');</script>"
	response.end
end if

mypic=rs("boysex")
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
if zg>=(4+int(DateDiff("d",zt(2),now())/10)) then
	Response.Write "<script Language=Javascript>alert('提示：您已经照顾"&(4+int(DateDiff("d",zt(2),now())/10))&"次了，请等5小时再操作！');</script>"
	response.end
end if
duyao=abate(myhua,name,1)
'名字|类别|生日|体力|攻击|心情|照顾数次|照顾时间
'金龍|金龍|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))+50)&"|"&zt(4)&"|"&zt(5)&"|"&(zg+1)&"|"&now()
conn.execute "update 用户 set w1='"&duyao&"',boy='"&temp&"' where  姓名='"&sjjh_name&"'"
conn.execute "update 用户 set boy='"&temp&"' where  姓名='"&zz&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：小孩喂养["&name&"]成功！');parent.f3.location.href = 'boy.asp';</script>"
%>
