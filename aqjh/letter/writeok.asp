<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<!--#include file="ffsj.asp"-->
<%
if Session("aqjh_inthechat")<>"1" then
	call subend("要想发送短信你必须进入聊天室。")
end if
if aqjh_jhdj<50 then
	call subend("提示：50级以上才可以发短信，你只有"&aqjh_jhdj&"级！")
end if
bt=trim(lcase(request.form("bt")))
name=left(trim(lcase(request.form("name"))),5)
txt=trim(lcase(request.form("txt")))
if name="" or txt="" or bt="" then
	call subend("提示：收件人、信件标题、信件内容均不能为空！")
end if
if name=aqjh_name then
	call subend("提示：收件人和发件人不能为同一人！")
end if
call charjc(name,"提示：名字中含有非法字符，请使用中文字符。")
if len(bt)>20 then
	call subend("信件标题最多为20个字，你的标题超出了范围！")
end if
if len(txt)>200 then
	call subend("信件内容最多为200个字，你的信件内容太长！")
end if
bt=jc(bt)
txt=jc(txt)
if bt="" then bt="无标题"
if txt="" then txt="无内容"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两,金币 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("金币")<1 then
	call myclose("提示：江湖使者替你送一次信收费1金币，你的钱不够。",0)
end if
rs.close
rs.open "select top 1 rq from letter where fjr='"&aqjh_name&"' order by rq desc",conn
if not(rs.eof or rs.bof) then
	sj=DateDiff("n",rs("rq"),now())
	if sj<5 and aqjh_grade<10 then
		s=2-sj
		call myclose("提示：信件2分钟发一次，你才发过信件，请等"&s&"分钟以后再发送！",0)
	end if
end if
rs.close
rs.open "select id from 用户 where 姓名='"&name&"'",conn
if rs.eof or rs.bof then
	call myclose("提示：名字为:"&name&"的收件人账号不存在！",0)
end if
yhid=rs("id")
rs.close
dim tmprs
tmprs=conn.execute("Select count(*) As 数量 from letter where yhid="&yhid)
letternumber=tmprs("数量")
set tmprs=nothing
if letternumber<10 or letternumber="" then
	conn.execute "insert into letter (yhid,bt,xj,rq,fjr,sjr) values ("&yhid&",'"&bt&"','"&txt&"',now(),'"&aqjh_name&"','"&name&"')"
else
	rs.open "select top 1 * from letter where yhid="&yhid&" order by rq",conn,1,3
	rs("yhid")=yhid
	rs("bt")=bt
	rs("xj")=txt
	rs("rq")=now()
	rs("fjr")=aqjh_name
	rs("sjr")=name
	rs("cs")=false
	rs("del")=false
	rs.update
	rs.close
end if
conn.execute "update 用户 set 金币=金币-1 where 姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
aqjh_roominfo=split(Application("aqjh_room"),";")
roomsn=ubound(aqjh_roominfo)-1
fx=0
for i=0 to roomsn
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(name)&" ")<>0 then
		nowinroom=i
		fx=1
		exit for
	end if
next
erase aqjh_roominfo
if fx=1 then
	says="<font color=red><b>【短消息】</b></font>"&aqjh_name&"给你发送了一条短信，快去查收一下呀！<a href=../letter/letter.asp target=_blank>单击这里查收短信</a>"
	jjj="<a href=javascript:parent.sw('[" & aqjh_name & "]'); target=f2><font color="&addwordcolor&">"& aqjh_name & "</font></a>"
	says=replace(says,aqjh_name,jjj)
	says=replace(says,name,jjk)
	says=replace(says,chr(39),"\'")
	says=replace(says,chr(34),"\"&chr(34))
	act="消息"
	towhoway=1
	towho=name
	addwordcolor="660099"
	saycolor="008888"
	addsays="对"
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
	addmsg saystr
end if
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<HTML><HEAD><TITLE>短消息发送成功</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../jhimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=../jhimg/back002.gif bgColor=#7c6244 leftMargin=0 oncontextmenu="return false;" text=#cccccc topMargin=0 marginheight="0" marginwidth="0">
<div align=center><br>
<table border="0" cellpadding="0" cellspacing="0" width="504">
<tr>
<td width="22"><img src="../jhimg/jiao_z_s.gif" width="21" height="21"></td>
<td width="468" background="../jhimg/line_s.gif"></td>
<td width="14"><img src="../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr>
<td width="22" background="../jhimg/line_z.gif">&nbsp;</td>
<td width="468" align=center> 
<TABLE bgColor=#675138 border=0 cellPadding=5 cellSpacing=1 width=521 height="195" style="border-left: 1 solid #816445; border-right: 1 solid #A98761; border-top: 1 solid #A98761">
<TBODY> 
<TR align=middle> 
<TD height="15" bgcolor="#333300" colspan="2" width="445"> <SPAN class=title><IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17> 
成功发送短消息 <IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17></SPAN> 
</TD>
</TR>
<TR> 
<TD height="90" bgcolor="#675138" width="445"> 
<div align="center">恭喜，您已成功将短消息发送给<%=name%><br>
<br>
<input type="button" value="关闭" onClick="window.close()" name="button">
<br>
</div>
</TD>
</TR>
</TBODY> 
</TABLE>
<!--#include file="bottom.htm"-->
</td>
<td width="14" background="../jhimg/line_y.gif">&nbsp;</td>
</tr>
<tr>
<td width="22"><img src="../jhimg/jiao_z_x.gif" width="22" height="21"></td>
<td width="468" background="../jhimg/line_x.gif"></td>
<td width="14"><img src="../jhimg/jiao_y_x.gif" width="21" height="21"></td>
</tr>
</table>
</div>
</BODY></HTML>
