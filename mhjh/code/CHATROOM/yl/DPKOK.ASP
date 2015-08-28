<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 

nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")

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
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
id=trim(request.querystring("id"))
fromname=trim(request.querystring("from1"))
to1=trim(request.querystring("to1"))
yn=trim(request.querystring("yn"))

if to1="大家" or to1=fromname then call ErrALT("你不可以选择大家或自已作为对手！")

if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
Response.End
end if
if nickname<>to1 then
  msg="人家没有邀请你呀！"
  abc=0
elseif yn=0 then
  msg="你拒绝人家的对赌了！"
  abc=1
  duidu="【拒绝】["&nickname&"]拒绝["& fromname &"]的打扑克牌请求!"
  duidu=duidu & "<br>.........打扑克难搞得紧，我还要研究一下..."
else
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	db="dpk.asp"
	'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
	'如果你的服务器支持jet4.0，请使用下载的连接方法，提高程序性能
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.open connstr 
	sql="select * from dpk where ufrom='"& nickname & "' or ufrom='"& fromname & "'"
	Set Rs=conn.Execute(sql)
	if not (rs.eof and rs.bof) then
		unpkname=rs("ufrom")
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		response.write "<script Language='Javascript'>alert('" & unpkname & "还有牌局没有结束，不能再次开局')</script>"
		abc=1
		duidu="【出错】["&nickname&"]不能接授["& fromname &"]的对赌请求!,因为["&unpkname&"]还有其他的牌局没有结束。"
		msg=unpkname & "还有牌局没有结束，不能再次开局！"
	else
		sql="select duz from dpk where uto='"& nickname & "$下注' and ufrom='"& fromname & "$下注' order by id Desc"
		Set Rs=conn.Execute(sql)
		If rs.eof and rs.bof Then
			rs.close
			conn.close
			set rs=nothing
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('错误：没有这个牌局！');}</script>"
			response.end
		end If
		xiazhuold=Rs(0)
		dim xia_fir,xiazhu,xia_sql,dppd,xia_class
		xia_fir=left(xiazhuold,1)
		xiazhu=mid(xiazhuold,2)
		xiazhu=clng(xiazhu)
		select case xia_fir
			case "1"
				xia_class="金币"
				xia_sql="金币"
			case "2"
				xia_class="银两"
				xia_sql="银两"
			case "3"
				xia_class="豆点"
				xia_sql="泡豆点数"
			case else
				conn.close
				set rs=nothing
				set conn=nothing
				call ErrALT("非法操作！")
		end select
		sql="delete * from dpk where uto='"& nickname & "$下注' or ufrom='"& fromname & "$下注' or ufrom='"& nickname & "$下注' or uto='"& fromname & "$下注'"
		Set Rs=conn.Execute(sql)
		dppd=0
		Set conn1=Server.CreateObject("ADODB.CONNECTION")
		Set rs1=Server.CreateObject("ADODB.RecordSet")
		conn1.open application("sjjh_usermdb")
		sql="select "& xia_sql &" from 用户 where 姓名='"& nickname &"'"
		set rs1=conn1.execute(sql)
		sjxiazhu=rs1(0)
		rs1.close
		if xiazhu>sjxiazhu then
			msg=nickname&"的赌金不够，不能开局。"
			duidu="【出错】["&nickname&"]不能接授["& fromname &"]的对赌请求!,因为["&nickname&"]的赌金不够。"
			dppd=1
		end if
		sql="select "& xia_sql &" from 用户 where 姓名='"& fromname &"'"
		set rs1=conn1.execute(sql)
		sjxiazhu=rs1(0)
		rs1.close
		if xiazhu>sjxiazhu then
			msg=fromname&"的赌金不够，不能开局。"
			duidu="【出错】["&nickname&"]不能接授["& fromname &"]的对赌请求!,因为["&fromname&"]的赌金不够。"
			dppd=1
		end if
		if dppd=0 then
				'先在双方身上去掉所下赌注
			sql="update 用户 set "& xia_sql &"="& xia_sql &"-"& xiazhu &" where 姓名='"& nickname &"' or 姓名='"& fromname &"'"
			set rs1=conn1.execute(sql)
		end if
		set rs1=nothing
		conn1.close
		set conn1=nothing
		abc=1
		if dppd=0 then
			for i= 1 to 18
				Randomize
				apk = Int(14 * Rnd+1)
				Randomize
				bbb= Int(4 * Rnd+1)
					if bbb=1 then bbb2="红"
				if bbb=2 then bbb2="黑"
				if bbb=3 then bbb2="梅"
				if bbb=4 then bbb2="方"
				dpk=bbb2 & apk
				dpk=convpk(dpk)
				if instr(frpk,dpk & "|")=0 then
					frpk=frpk & dpk & "|"
				else 
					i=i-1
				end if
			next
			for i= 1 to 18
				Randomize
				apk = Int(14 * Rnd+1)
				Randomize
				bbb= Int(4 * Rnd+1)
				if bbb=1 then bbb2="红"
				if bbb=2 then bbb2="黑"
				if bbb=3 then bbb2="梅"
				if bbb=4 then bbb2="方"
				dpk=bbb2 & apk
				dpk=convpk(dpk)
				if instr(frpk,dpk & "|")=0 and instr(topk,dpk & "|")=0 then
					topk=topk & dpk & "|"
				else 
					i=i-1
				end if
			next
			sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& nickname & "','"& fromname & "','"& topk & "',false," & xiazhuold & ",90)"
			Set Rs=conn.Execute(sql)
			sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& fromname & "','"& nickname & "','"& frpk & "',true," & xiazhuold & ",90)"
			Set Rs=conn.Execute(sql)
			duidu="<b><font color=green>["&nickname&"]</font></b>同意打扑克,<b><font color=green>["&application("sjjh_automanname")&"]</font></b>开始分牌......<br>"
			duidu=duidu & "<b><font color=green>["&nickname&"]</font></b>抓了18张牌<img src='f2/dpk/1.GIF'>"
			duidu=duidu & "<input type=button value='洗牌' onclick=""javascript:window.open('f2/dpk-xp.asp?name="&nickname&"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dpkname""><img src='f2/dpk/2.GIF'>"
			duidu=duidu & " <b><font color=green>["
			duidu=duidu & fromname&"]</font></b>抓了18张牌<img src='f2/dpk/2.GIF'><input type=button value='洗牌' onclick=""javascript:window.open('f2/dpk-xp.asp?name="
			duidu=duidu & fromname &"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dpkname""><img src='f2/dpk/1.GIF'>"
			duidu=duidu & "<br>.....根据牌局规定，请[" & nickname & "]先发牌。<br>"
			abc=1
			msg="你同意跟对方打扑克了，请稍候！[江湖博士]正在帮你们洗牌！"
			conn.close
			set rs=nothing
			set conn=nothing
		end if
	end if
end if

if abc=1 then
says="<font color=#ff0000>【打扑克】</font>："&duidu			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
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
end if
%>
<html>
<head>
<style>
a:link {text-decoration:none; font-size:14;color:#000000}
a:hover {text-decoration:underline;font-size:14; color:#000000; background:#ffffff}
a:visited {text-decoration:none;font-size:14; color:#000000}
td {font-size:9pt; color:#ff0000; line-height:16pt}
</style>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>大庆娱乐江湖-打扑克</title>
</head>
<body bgcolor=#FFFFFF background="bg1.gif" >
<tr>
<td> <font color="red"> </font><br>
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#FFFFFF">
<tr> 
<td height="115"> 
          <table border="0" bgcolor="#0000FF" cellspacing="0" cellpadding="2" width="361">
            <tr> 
<td width="324" height="9"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">八恒江湖提示</font><font color="red" size=2> 
</font></td>
<td width="29" height="9"> 
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr> 
<td width="16"><b><a href="<%if id="200" then%>javascript:window.close();<%else%>javascript:history.go(-1)<%end if%>" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="返回"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="359" cellpadding="4">
<tr> 
<td width="59" align="center" valign="top" height="29"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
<td width="278" height="29"> <font color="red" size=2> 
<%=msg%>
</font></td>
</tr>
<tr> 
<td colspan="2" align="center" valign="top" height="58"> 
<input type=button value='关闭' onClick="javascript:window.close()" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button223">
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<%
function convpk(dpk)
if dpk="红14" then dpk="大王"
if dpk="方14" then dpk="大王"
if dpk="梅14" then dpk="小王"
if dpk="黑14" then dpk="小王"
dpk=replace(dpk,"13","K")
dpk=replace(dpk,"12","Q")
dpk=replace(dpk,"11","J")
dpk=replace(dpk,"10","十")
dpk=replace(dpk,"1","A")
dpk=replace(dpk,"十","10")
convpk=dpk
end function 
if to1=jiutian_dabusi then
%>
<script language="JavaScript">window.close();</script>
<%end if%>
</html>


