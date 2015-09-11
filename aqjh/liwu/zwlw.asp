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
	weekdate=weekday(date())
if weekdate<>6 then
	Response.Write "<script Language=Javascript>alert('周五礼物只能周五领取今天是礼拜"&weekdate-1&"！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,mvalue,w7,等级,会员等级,转生 FROM 用户 WHERE 姓名='" & aqjh_name & "'",conn
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
if rs("mvalue")<=20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('月积分在20000以上的才可以领取物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
myid=rs("id")
wpdate=rs("w7")
hydj=rs("会员等级")
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("sdlw.mdb")
rs1.open "select * from 周五礼物 where yhid="&myid,conn1
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
randomize timer
r=int(rnd*5)+1
select case r
	case 1 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=1000*hydj
                p=7000*hydj
		temp=add(wpdate,"美满生活",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&m&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="美满生活"&n&"个,金"&m&",木"&k&",水"&k&",火"&m&",土"&k&",法力"&l&",轻功"&l&",武功"&p&""
	case 2 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=1000*hydj
                p=80000*hydj
		temp=add(wpdate,"求爱玫瑰",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&m&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="求爱玫瑰"&n&"个,金"&m&",木"&k&",水"&k&",火"&m&",土"&k&",法力"&l&",轻功"&l&",武功"&p&""	
	case 3 
		n=999*hydj
                m=400*hydj
                k=600*hydj
                l=800*hydj
                p=6000*hydj
		temp=add(wpdate,"鹊桥相会",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&l&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="鹊桥相会"&n&"个,金"&m&",木"&k&",水"&k&",火"&l&",土"&k&",法力"&l&",轻功"&l&",武功"&p&""
       case 4 
		n=999*hydj
                m=700*hydj
                k=600*hydj
                l=500*hydj
                p=6000*hydj
		temp=add(wpdate,"玫瑰衷情",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&l&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="玫瑰衷情"&n&"个,金"&m&",木"&k&",水"&k&",火"&l&",土"&k&",法力"&l&",轻功"&m&",武功"&p&""
	case 5 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=900*hydj
                p=7000*hydj
		temp=add(wpdate,"跳舞的玫瑰",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&k&",土=土+"&l&",法力=法力+"&k&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="跳舞的玫瑰"&n&"个,金"&m&",木"&k&",水"&k&",火"&k&",土"&k&",法力"&k&",轻功"&l&",武功"&p&""	
	case 6 
		n=999*hydj
                m=600*hydj
                k=500*hydj
                l=700*hydj
                p=5000*hydj
		temp=add(wpdate,"想飞的玫瑰",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&l&",土=土+"&m&",法力=法力+"&k&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="想飞的玫瑰"&n&"个,金"&m&",木"&k&",水"&k&",火"&l&",土"&m&",法力"&k&",轻功"&l&",武功"&p&""
end select
set rs=conn.execute(sql)
conn1.execute "insert into  周五礼物(yhid,姓名,礼物) values ('"&myid&"','"&aqjh_name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
conn1.close
set conn1=nothing
says="<font color=red><b>【周五礼物】</b></font>"&aqjh_name&"从站长那里得到周五礼物"&lw&""

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
