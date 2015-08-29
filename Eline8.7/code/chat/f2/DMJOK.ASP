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
id=LCase(trim(request.querystring("id")))
fromname=trim(request.querystring("from1"))
to1=trim(request.querystring("to1"))
yn=trim(request.querystring("yn"))

if to1="大家" or to1=fromname then call ErrALT("你不可以选择大家或自已作为对手！")
if sjjh_jhdj<15 then
	call mess ("被邀打扑克需要15级才可以操作哦^_^",1)
end if
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
	duidu="【拒绝】["&nickname&"]拒绝["& fromname &"]的搓麻将请求!"
	duidu=duidu & "<br>.........搓麻将难搞得紧，我还要研究一下...."
else
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	db="eDMJ.ASP"
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.open connstr
	sql="select * from dmj where ufrom='"& nickname & "' or ufrom='"& fromname & "' or uto='"& nickname &"' or uto='"& fromname &"'"
	Set Rs=conn.Execute(sql)
	if not (rs.eof and rs.bof) then
	  if rs("ufrom")=nickname or rs("ufrom")=fromname then
		  unpkname=rs("ufrom")
	  else
	  	  unpkname=rs("uto")
	  end if
	  rs.close
	  set rs=nothing
	  conn.close
	  set conn=nothing
	  response.write "<script Language='Javascript'>alert('" & unpkname & "还有牌局没有结束，不能再次开局')</script>"
	  abc=1
	  duidu="【出错】["&nickname&"]不能接授["& fromname &"]的对赌请求!,因为["&unpkname&"]还有其他的牌局没有结束。"
	  msg=unpkname&"还有牌局没有结束，不能再次开局！"
	else
	  sql="select duz from dmj where uto='"& nickname & "$下注' and ufrom='"& fromname & "$下注' order by id Desc"
	  Set Rs=conn.Execute(sql)
	  If rs.eof and rs.bof Then
	  		Response.Write "<script language=JavaScript>{alert('错误：没有这个赌局！');}</script>"
			rs.close
			conn.close
			set rs=nothing
			set conn=nothing
			response.end
	  end If
	  xiazhu=Rs(0)
	  sql="delete * from dmj where uto='"& nickname & "$下注' or ufrom='"& fromname & "$下注' or ufrom='"& nickname & "$下注' or uto='"& fromname & "$下注'"
	  Set Rs=conn.Execute(sql)
	  dim xia_class,xia_fir,xia_sql,duidupd
		xia_fir=left(xiazhu,1)
		xiazhu1=mid(xiazhu,2)
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
				call ErrALT("非法操作！")
		end select
		duidupd=0
		Set conn1=Server.CreateObject("ADODB.CONNECTION")
		Set rs1=Server.CreateObject("ADODB.RecordSet")
		conn1.open application("sjjh_usermdb")
		sql="select " & xia_sql & " from 用户 where 姓名='"& nickname &"'"
		set rs1=conn1.execute(sql)
		sjxiazu=rs1(0)
		rs1.close
		if sjxiazu<(clng(xiazhu1)*2) then
			msg=nickname&"的赌注"&xia_class&"不足，不能打麻将。要有所下赌注两倍以上的资金才能打麻将。"
			abc=1
			duidu=nickname&"的赌注"&xia_class&"不足，不能打麻将。要有所下赌注两倍以上的资金才能打麻将。"
			duidupd=1
		end if
		myxiazu=sjxiazu-clng(xiazhu1)*2
		sql="select " & xia_sql & " from 用户 where 姓名='"& fromname &"'"
		set rs1=conn1.execute(sql)
		sjxiazu=rs1(0)
		rs1.close
		if sjxiazu<(clng(xiazhu1)*2) then
			msg=fromname&"的赌注"&xia_class&"不足，不能打麻将。要有所下赌注两倍以上的资金才能打麻将。"
			abc=1
			duidu=fromname&"的赌注"&xia_class&"不足，不能打麻将。要有所下赌注两倍以上的资金才能打麻将。"
			duidupd=1
		end if
		toxiazu=sjxiazu-clng(xiazhu1)*2
		if duidupd=0 then	'扣除所下赌金的2倍，待决出胜负后再返还
			sql="update 用户 set "& xia_sql &"="& myxiazu &" where 姓名='"& nickname &"'"
			set rs1=conn1.execute(sql)
			sql="update 用户 set "& xia_sql &"="& toxiazu &" where 姓名='"& fromname &"'"
			set rs1=conn1.execute(sql)
		else
			sql="delete * from mjInfo where 庄家='" & nickname & "' or 庄家='" & fromname & "' or 对手='" & nickname & "' or 对手='" & fromname & "'"
			Set Rs=conn.Execute(sql)
		end if
		set rs1=nothing
		conn1.close
		set conn1=nothing
		if duidupd=0 then
			dim Allmjp
			Allmjp="1万|2万|3万|4万|5万|6万|7万|8万|9万|1筒|2筒|3筒|4筒|5筒|6筒|7筒|8筒|9筒|1索|2索|3索|4索|5索|6索|7索|8索|9索|东风|南风|西风|北风|红中|白板|发财|"
			Allmjp=Allmjp & Allmjp & Allmjp & Allmjp
			mjpArr=split(Allmjp,"|")
			Randomize
			for i= 1 to 14
				mjpArr=split(Allmjp,"|")
				Mjx= Int(ubound(mjpArr)* Rnd)
				strMjx=mjpArr(Mjx)
				Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
				myMj=myMj & strMjx & "|"
			next
			Randomize
			for i= 1 to 13
				mjpArr=split(Allmjp,"|")
				Mjx= Int(ubound(mjpArr) * Rnd)
				strMjx=mjpArr(Mjx)
				Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
				youMj=youMj & strMjx & "|"
			next
			Allmjp_Akk=""
			Randomize
			for i= 1 to 109
				mjpArr = Split(Allmjp, "|")
				Mjx= Int(UBound(mjpArr) * Rnd)
				strMjx=mjpArr(Mjx)
				Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
				Allmjp_Akk=Allmjp_Akk & strMjx & "|"
			next
			sql="delete * from mjInfo where 庄家='" & nickname & "' or 庄家='" & fromname & "' or 对手='" & nickname & "' or 对手='" & fromname & "'"
			Set Rs=conn.Execute(sql)
			sql="insert into mjInfo(庄家,对手,麻将) values ('"& nickname & "','"& fromname & "','"& Allmjp_Akk & "')"
			Set Rs=conn.Execute(sql)
			sql="select id from mjInfo where 庄家='"& nickname & "'"
			Set Rs=conn.Execute(sql)
			mjID=Rs("id")
			sql="insert into dmj(ufrom,uto,Mymj,duz,isMy,isGet,isFp,mjID) values ('"& nickname & "','"& fromname & "','"& Mymj & "'," & xiazhu & ",true,false,true," & mjID & ")"
			Set Rs=conn.Execute(sql)
			sql="insert into dmj(ufrom,uto,Mymj,duz,isMy,isGet,isFp,mjID) values ('"& fromname & "','"& nickname & "','"& youMj & "'," & xiazhu & ",false,false,false," & mjID & ")"
			Set Rs=conn.Execute(sql)
			duidu="<b><font color=green>["&nickname&"]</font></b>同意搓麻将,<b><font color=green>["&application("sjjh_automanname")&"]</font></b>洗牌......唏哩哗啦一阵暴响<br>"
			duidu=duidu & "<b><font color=green>["&nickname&"]</font></b>摸了14张牌<img src='f2/sz.gif'>"
			duidu=duidu & "<input type=button value='洗牌' onclick=""javascript:window.open('f2/dmj-xp.asp?name="&nickname&"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dmjname""><img src='f2/sz.gif'>"
			duidu=duidu & " <b><font color=green>["
			duidu=duidu & fromname&"]</font></b>摸了13张牌<img src='f2/sz.gif'><input type=button value='洗牌' onclick=""javascript:window.open('f2/dmj-xp.asp?name="
			duidu=duidu & fromname &"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dmjname""><img src='f2/sz.gif'>"
			duidu=duidu & "<br>.....现在[" & nickname & "]是庄家，请先打牌。<br>"
			abc=1 
			msg="你同意跟对方搓麻将了，请稍候！[江湖博士]正在帮你们洗牌！"
		end if
		set conn=nothing
		set rs=nothing
	end if
end if
if abc=1 then
says="<font color=#ff0000>【搓麻将】</font>："&duidu			'聊天数据
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
<title>E线江湖-打麻将</title>
</head>
<body bgcolor=#FFFFFF background="bg1.gif" >
<tr>
<td> <font color="red"> </font><br>
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#FFFFFF">
<tr> 
<td height="115"> 
          <table border="0" bgcolor="#0000FF" cellspacing="0" cellpadding="2" width="361">
            <tr> 
<td width="324" height="9"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">E线江湖提示</font><font color="red" size=2> 
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
function strMjp(cmj)
dim mj2
dim mjr
dim mjL
mjr=right(cmj,1)
if mjr=0 or cmj>40 then
mj2=replace(cmj,"10","东风")
mj2=replace(cmj,"20","南风")
mj2=replace(cmj,"30","西风")
mj2=replace(cmj,"40","北风")
mj2=replace(mj2,"41","红中")
mj2=replace(mj2,"42","白板")
mj2=replace(mj2,"43","发财")
else
mjL=Left(cmj,1)
mjL=replace(mjL,"1","索")
mj2=replace(mjL,"2","筒")
mjL=replace(mjL,"3","万")
mj2=mjr & mjL
end if
strMjp=mj2
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL left(cmj,1)
if isNumeric(mjL) then mj2=right(cmj,1) & mjL
mj2=replace(mj2,"索","1")
mj2=replace(mj2,"筒","2")
mj2=replace(mj2,"万","3")
mj2=replace(mj2,"风","0")
mj2=replace(mj2,"东","1")
mj2=replace(mj2,"南","2")
mj2=replace(mj2,"西","3")
mj2=replace(mj2,"北","4")
mj2=replace(mj2,"红中","41")
mj2=replace(mj2,"白板","42")
mj2=replace(mj2,"发财","43")
intMjp=int(mj2)
end function
if to1="大家" or to1=fromname then
%>
<script language="JavaScript">window.close();</script>
<%end if%>
</html>
