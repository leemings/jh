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
rs.open "SELECT id,mvalue,w5,等级,会员等级,转生 FROM 用户 WHERE 姓名='" & aqjh_name & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你现在还不是我们江湖的人吧！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
If rs("等级")<100 and rs("转生")<2 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('转生2次，并且等级达到100级才能领取！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("mvalue")<=50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('月积分在50000以上的才可以领取物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
myid=rs("id")
wpdate=rs("w5")
hydj=rs("会员等级")
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("sdlw.mdb")
rs1.open "select * from 站长礼物 where yhid="&myid,conn1
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已经领过礼物了，或者站长还没有给大家发礼物！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if




rs1.close
lw=""
randomize timer
r=int(rnd*13)+1
select case r
	case 1 '金币
		n=int(rnd*10)+2
		sql="update 用户 set 金币=金币+"&n&" where id="&myid
		lw="金币"&n&"个"
	case 2 '金卡
		n=int(rnd*5)+4
		sql="update 用户 set 会员金卡=会员金卡+"&n&" where id="&myid
		lw="会员金卡"&n&"个"
	case 3 '宝物卡
		n=int(rnd*2)+1
		temp=add(wpdate,"宝物卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="宝物卡"&n&"个"
	case 4 '真爱卡
		n=int(rnd*2)+1
		temp=add(wpdate,"真爱卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="真爱卡"&n&"个"
	case 5 '养猪卡
		n=int(rnd*2)+1
		temp=add(wpdate,"养猪卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="养猪卡"&n&"个"
	case 6 '补血卡
		n=int(rnd*1)+1
		temp=add(wpdate,"补血卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="补血卡"&n&"个"
	case 7 '财神卡
		n=int(rnd*2)+1
		temp=add(wpdate,"财神卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="财神卡"&n&"个"
	case 8 '种花卡
		n=int(rnd*1)+1
		temp=add(wpdate,"种花卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="种花卡"&n&"个"
	case 9 '送神卡
		n=int(rnd*2)+1
		temp=add(wpdate,"送神卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="送神卡"&n&"个"
	case 10 '加点卡
		n=int(rnd*2)+1
		temp=add(wpdate,"加点卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="加点卡"&n&"个"
	case 11 '天之卡
		n=int(rnd*2)+1
		temp=add(wpdate,"天之卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="天之卡"&n&"个"
	case 12 '练功卡
		n=int(rnd*2)+1
		temp=add(wpdate,"练功卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="练功卡"&n&"个"
	case 13 '战神卡
		n=int(rnd*2)+1
		temp=add(wpdate,"战神卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="战神卡"&n&"个"
end select
set rs=conn.execute(sql)
conn1.execute "insert into  站长礼物(yhid,姓名,礼物) values ('"&myid&"','"&aqjh_name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
conn1.close
set conn1=nothing
says="<font color=red><b>【站长礼物】</b></font>"&aqjh_name&"从站长那里得到站长礼物"&lw&""

call showchat(says)

%>
<html>
<head>
<title>礼物发放</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../chat/bg.gif">
<div align="center">
<table width="141" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="141" align="center"><b><font color="#FFFF00">恭喜你从<%=application("aqjh_user")%>那里得到了</font><br><font color="#FF0000"><%=lw%></font></b></td>
  </tr>
</table>
</div>
</body>
</html>
</body>
</html>
