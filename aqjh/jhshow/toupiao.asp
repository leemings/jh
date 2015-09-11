<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
ndate=Day(date())
if ndate<21 or ndate>27 then
		Response.Write "<script Language=Javascript>alert('提示：投票时间为每月21到27号，现在时间不可以进行！');window.close();</script>"
		response.end
end if
id=clng(Request("id"))
if Session("aqjh_jhdj")<20 then
	Response.Write "<script Language=Javascript>alert('提示：你是江湖新手，不够20级不能参与投票！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 性别,配偶,情人,介绍人,师傅 from [用户] where 姓名='" & Session("aqjh_name") & "'",conn,1,1
mysex=rs("性别")
po=rs("配偶")
qr=rs("情人")
jsr=rs("介绍人")
sf=rs("师傅")
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,1
if instr(rs("p_投票人"),","& aqjh_name &"|")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：江湖规定，每人只可以投一票，你已经投过了！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.Open "select * from sai where s_id="& id, conn, 1,3
if mysex=rs("s_性别") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：江湖规定只可以异性投票，即男对女进行投票，女对男进行投票！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
toname=rs("s_姓名")
if po=toname or qr=toname or jsr=toname or sf=toname then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：配偶、情人、介绍人、师傅人不可以进行投票！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs("s_票数")=rs("s_票数")+1
rs.update
rs.close
rs.Open "select * from pool", conn, 1,3
rs("p_投票人")=rs("p_投票人") & toname & "," & aqjh_name & "|"
rs("p_投票人数")=rs("p_投票人数")+1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两 from [用户] where 姓名='" & Session("aqjh_name") & "'",conn,1,3
rs("银两")=rs("银两")+10000
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【消息】##在江湖秀大赛中投出保贵的一票，系统奖励1万两！江湖秀大赛现在进行中，要进行评选的快去。。。<font class=t>(" & time() & ")</font></font>"
act="消息"
towho="大家"
addwordcolor="660099"
saycolor="008888"
call addsay(Session("inroom"),aqjh_name,"对",towho,addwordcolor,saycolor,says,act,0)
Response.Write "<script Language=Javascript>alert('感谢您投出的这一票，系统自动为您加上一万两作为奖励！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>