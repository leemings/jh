<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<!--#include file="../showchat.asp"-->

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,mvalue,w5,等级,会员等级 FROM 用户 WHERE 姓名='" & aqjh_name & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你现在还不是我们江湖的人吧！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("等级")<>55 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('为防止作弊，只有等级在55级的时候才可以领取物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("mvalue")<=8000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('月积分在8000以上的才可以领取物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
myid=rs("id")
wpdate=rs("w5")
hydj=rs("会员等级")
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("sdlw.mdb")
rs1.open "select * from 新人礼物 where yhid="&myid,conn1
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已经领过礼物了，礼物不多，还是给其他人也留点吧！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs1.close
lw=""
		jb=200
		sql="update 用户 set 金币=金币+"&jb&" where id="&myid
		lw0="金币"&jb&"个"

		jk=2
		sql1="update 用户 set 会员金卡=会员金卡+"&jk&" where id="&myid
		lw1="会员金卡"&jk&"个"

		kp=2
		temp=add(wpdate,"福神卡",kp)
		sql2="update 用户 set w5='"&temp&"' where id="&myid
		lw="福神卡"&kp&"个"
set rs=conn.execute(sql)
set rs=conn.execute(sql1)
set rs=conn.execute(sql2)
conn1.execute "insert into  新人礼物(yhid,姓名,礼物) values ('"&myid&"','"&aqjh_name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
conn1.close
set conn1=nothing
says="<font color=red><b>【江湖消息】</b></font>"&aqjh_name&"从站长那里得到"&lw0&""&lw1&""&lw&""

call showchat(says)



%>
<html>
<head>
<title>礼物发放</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body text="#000000" background="../chat/bg.gif">
<div align="center">
<table width="106" border="0" cellspacing="0" cellpadding="0" bordercolor="#0A246A" bgcolor="#808000">
  <tr>
    <td width="26">领取成功</td>
    <td width="80" align="center"><b><font color="#FFFF00">恭喜你从<%=application("aqjh_user")%>那里得到了</font><br><font color="#FF0000"><%=lw0%><%=lw1%><%=lw%></font></b></td>
  </tr>
</table>
</div>
</body>
</html>
</body>
</html>
