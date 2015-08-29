<!--#include file="CONN.ASP"-->
<!--#include file="z_conncaipiao.ASP"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_plus_check.asp"-->
<%
response.buffer=true
dim MinJiang,JiaGe,AMaxJiang1,AMaxJiang2,AMaxJiang3,AMaxJiang4,CP_Msg
dim CP_Qi,CP_MaxJiang,CP_TouZhuRenShu,CP_TouZhuE,CP_OpenDateTime,UserMoney
' 奖池最小奖金数量，低于此数量自动补足(需要新一期才可生效)
MinJiang	=	20000
' 单注价格
JiaGe		=	5
' 单注最高奖金
AMaxJiang1 =	15000
AMaxJiang2 =	1500
AMaxJiang3 =	150
AMaxJiang4 =	15


Stats="彩票中心"
call nav()
call head_var(3,0,"","")
response.write CP_Msg
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入彩票中心的权限，请先<a href=login.asp><font color=blue>登录</font></a>或者同管理员联系。"
	call dvbbs_error()
else
set rs=server.createobject("adodb.recordset")

rs.open "select * from [user] where userid="&userid,conn,1,3
usermoney=rs("userwealth")
rs.close

dim rs3,rs1
set rs3=server.createobject("adodb.recordset")
rs3.open "select Top 1 * From CP_Record where OpenDateTime>=#"&year(now)&"-"&month(now)&"-"&day(now)&" "&hour(now)&":"&minute(now)&":"&second(now)&"# order by id desc",conncaipiao,1,3
if rs3.eof or rs3.bof then
	set rs1=server.createobject("adodb.recordset")
	rs1.open "select Top 1 * From CP_Record order by id desc",conncaipiao,1,3
	if rs1.eof or rs1.bof then
		CP_MaxJiang=MinJiang
	Else
		call KiaJiang(RS1("ID"))
		CP_MaxJiang=RS1("YuE")
	End if
	if CP_MaxJiang<MinJiang then
		CP_MaxJiang=CP_MaxJiang+MinJiang
	end if
	Rs1.close
	set rs1=nothing
	rs3.addnew
	Randomize
	rs3("Number")=int((8999*rnd)+1000)
	rs3("OpenDateTime")=date()+1&" 20:00:00"
	rs3("YuE")=CP_MaxJiang
	rs3.update
else
	set rs1=server.createobject("adodb.recordset")
	rs1.open "select Top 1 * From CP_user order by id desc",conncaipiao,1,3
	if not( rs1.bof and rs1.eof )then
		if rs1("Qi")<rs3("ID") then
			Randomize
			rs3("Number")=int((8999*rnd)+1000)
			rs3("OpenDateTime")=date()+1&" 20:00:00"
			rs3.update
		end if 
	end if			
end if
CP_Qi=rs3("id")
CP_OpenDateTime=rs3("OpenDateTime")
CP_MaxJiang=rs3("YuE")
rs3.close
set rs3=nothing
if request("job")<>"" then 
	call savepiao
	response.redirect scriptname
end if
CP_TouZhuRenShu=0
CP_TouZhuE=0

rs.open "select count(id) From CP_User Where Qi="&CP_Qi,conncaipiao,1,3
CP_TouZhuRenShu=rs(0)
rs.close
CP_MaxJiang=CP_MaxJiang+CP_TouZhuRenShu*JiaGe
CP_TouZhuE=CP_TouZhuRenShu*JiaGe
%>

<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th><%=membername%>,欢迎光临彩票中心</th></tr>
<tr><td class=tablebody2>
<br>
<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style="word-break:break-all;">
<tr>
<TD vAlign=top class=tablebody1>
&nbsp;欢迎光临彩票中心！<br>
&nbsp;您现在就可以购买第 <font color=#FF0000><b><%=CP_Qi%></b></font> 期彩票，本期总共累计奖额 <font color=#FF0000><b><%=CP_MaxJiang%></b></font> 元。 <br>
&nbsp;单注最高奖额 <font color=#0000FF><b>一等奖</b></font>：<font color=#FF0000><b><%=AMaxJiang1%></b></font>元，<font color=#0000FF><b>二等奖</b></font>：<font color=#FF0000><b><%=AMaxJiang2%></b></font>元，<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=#0000FF><b>三等奖</b></font>：<font color=#FF0000><b><%=AMaxJiang3%></b></font>元，<font color=#0000FF><b>四等奖</b></font>：<font color=#FF0000><b><%=AMaxJiang4%></b></font>元。<br>
&nbsp;本期共有投注 <font color=#FF0000><b><%=CP_TouZhuRenShu%></b></font> 注，总投注额 <font color=#FF0000><b><%=CP_TouZhuE%></b></font> 元，于 <font color=#FF0000><b><%=CP_OpenDateTime%></b></font> 开奖。<br>
</td>
<td vAlign=top class=tablebody1 rowspan=2>
<B>历史投注:</b><br>
<%
	rs.open "select top 10 * from cp_user where user="&userid&" order by qi desc,id desc",conncaipiao,1,3
	i=0
	do while ( not rs.eof ) and i<10 
%>
第<%=rs("Qi")%>期:<%=right("0000"&rs("Haoma"),4)%><br>
<%
	 	rs.movenext
		i=i+1
	loop
	rs.close
%>
</td>
</tr>
<tr>
<TD vAlign=top class=tablebody1 align="center">
每注彩票 <font color=red><%=JiaGe%></font> 元，您现在共有 <font color=red><%=UserMoney%></font> 元。<br>您需要购买什么号码？
<br><br>
<form action="z_CaiPiao.asp" method=Get name=BuyCaiPiao>
<input type="hidden" name="job" value="save">
<select name="CP1">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<select name="CP2">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<select name="CP3">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<select name="CP4">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<input type=submit value="投注">
<input type="button" value="机选" onclick="SelRandom()">
<script>
function SelRandom() {
	BuyCaiPiao.CP1.options[Math.random()*10].selected=true;
	BuyCaiPiao.CP2.options[Math.random()*10].selected=true;
	BuyCaiPiao.CP3.options[Math.random()*10].selected=true;
	BuyCaiPiao.CP4.options[Math.random()*10].selected=true;
}
SelRandom()
</script>
</form>
<%=CP_Msg%>
</td></tr>
</table>
<br>
</td></tr></table>
<br>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th>彩票中心---本期投注情况</th></tr>
<tr><td class=tablebody2>
<br>
	<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style="word-break:break-all;">
		<tr>
			<TD class=tablebody2 align=center height=20>用户名</td>
			<TD class=tablebody2 align=center>投注数</td>
			<TD class=tablebody2 align=center>投注金额</td>  
		</tr>
<%
	rs.open "SELECT Count(UserName) AS UserNameOfCount,UserName FROM CP_User where qi="&CP_Qi&" GROUP BY UserName ORDER BY Count(UserName) DESC ",conncaipiao,1,3
	if rs.eof and rs.bof then
	%>
		<tr><td colspan=3 class=tablebody1 align=center height=20>本期还没有人投注</td></tr>
	<%
	else
		do while not rs.eof
	%>		
			<tr>
				<TD class=tablebody1 align=center height=20><%=rs(1)%></td>
				<TD class=tablebody1 align=center><%=rs(0)%></td>
				<TD class=tablebody1 align=center><%=rs(0)*JiaGe%></td>
			</tr>
	<%
			rs.movenext
		loop
	end if
	rs.close
%>			
	</table>
<br>
</td></tr></table>
<br>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th>彩票中心---最近10期中奖情况</th></tr>
<tr><td class=tablebody2>
<br>
	<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style="word-break:break-all;">
		<tr>
			<TD class=tablebody2 align=center height=20>期  数</td>
			<TD class=tablebody2 align=center>中奖号码</td>
			<TD class=tablebody2 align=center>总投注数</td>
			<TD class=tablebody2 align=center>总投注金额</td>			
			<TD class=tablebody2 align=center>奖池总金额</td>  
			<TD class=tablebody2 align=center>中奖注数</td> 
			<TD class=tablebody2 align=center>兑出奖金</td>
			<TD class=tablebody2 align=center>奖池剩余奖金</td> 			
		</tr>
<%
    set	rs=conncaipiao.execute("SELECT top 10 id,ZhongJiangHaoMa,ZongJiTouZhu,ZongJiTouZhujine,ZongJinE,ZhongJiangZhuShu,JiangJin,ShengY FROM qingkuang order by id DESC")
	if rs.eof and rs.bof then
	%>
		<tr><td colspan=8 class=tablebody1 align=center height=20>还没有历史中奖情况</td></tr>
	<%
	else
		do while not rs.eof
	%>		
			<tr>
			<TD class=tablebody1 align=center height=20><%=rs(0)%></td>
			<TD class=tablebody1 align=center><%=rs(1)%></td> 
			<TD class=tablebody1 align=center><%=rs(2)%></td>  
			<TD class=tablebody1 align=center><%=rs(3)%></td> 
			<TD class=tablebody1 align=center><%=rs(4)%></td> 
			<TD class=tablebody1 align=center><%=rs(5)%></td>
			<TD class=tablebody1 align=center><%=rs(6)%></td> 
			<TD class=tablebody1 align=center><%=rs(7)%></td> 
			</tr>
	<%
			rs.movenext
		loop
	end if
	rs.close
	set rs=nothing
%>			
	</table>
<br>
</td></tr></table>
<%
end if
conncaipiao.close
set conncaipiao=nothing
call activeonline()
call footer()

' 以下无需修改
sub KiaJiang(ID)
	dim ZongJiTouZhu,ZongJinE,ZhongJiangHaoMa,ZhongJiangZhuShu1,ZhongJiangZhuShu2,ZhongJiangZhuShu3,ZhongJiangZhuShu4,JiangJin,Content,ShengY
	dim ZhongJiangUser,Delta  '中奖名单   我来了添加
	' 获得总金额
	rs.open "select * from CP_Record where ID="&id,conncaipiao,1,3
	ZongJinE=rs("YuE")
	ZhongJiangHaoMa=rs("Number")
	rs.close
	rs.open "select count(id) from CP_User where Qi="&ID,conncaipiao,1,3
	ZongJiTouZhu=rs(0)
	rs.close
	ZongJinE=ZongJinE+ZongJiTouZhu*JiaGe

	' 获得一等奖注数
	rs.open "select count(ID) from CP_User where Qi="&id&" and HaoMa="&ZhongJiangHaoMa&"",conncaipiao,1,3    
	ZhongJiangZhuShu1=rs(0)  
	rs.close
	' 获得二等奖注数
	rs.open "select count(ID) from CP_User where Qi="&id&" and HaoMa<>"&ZhongJiangHaoMa&" and (HaoMa mod 1000)="&(ZhongJiangHaoMa mod 1000)&"",conncaipiao,1,3    
	ZhongJiangZhuShu2=rs(0)
	rs.close
	' 获得三等奖注数
	rs.open "select count(ID) from CP_User where Qi="&id&" and (HaoMa mod 1000)<>"&int(ZhongJiangHaoMa mod 1000)&" and (HaoMa mod 100)="&int(ZhongJiangHaoMa mod 100)&"",conncaipiao,1,3
	ZhongJiangZhuShu3=rs(0)  
	rs.close
	' 获得四等奖注数
	rs.open "select count(ID) from CP_User where Qi="&id&" and (HaoMa mod 100)<>"&(ZhongJiangHaoMa mod 100)&" and (HaoMa mod 10)="&(ZhongJiangHaoMa mod 10)&"",conncaipiao,1,3    
	ZhongJiangZhuShu4=rs(0)  
	rs.close
	' 总奖金
	JiangJin=ZhongJiangZhuShu1*AMaxJiang1+ZhongJiangZhuShu2*AMaxJiang2+ZhongJiangZhuShu3*AMaxJiang3+ZhongJiangZhuShu4*AMaxJiang4
	' 兑奖率
	if JiangJin>ZongJinE then
		Delta=ZongJinE/JiangJin
	else
		Delta=1
	end if
	' 支付奖金
	JiangJin=ZhongJiangZhuShu1*int(AMaxJiang1*Delta)+ZhongJiangZhuShu2*int(AMaxJiang2*Delta)+ _
		ZhongJiangZhuShu3*int(AMaxJiang3*Delta)+ZhongJiangZhuShu4*int(AMaxJiang4*Delta)
	' 剩余奖金
	ShengY=ZongJinE-JiangJin
	rs.open "select * from CP_Record where ID="&id,conncaipiao,1,3
	rs("YuE")=ShengY
	rs.update
	rs.close
	if ZhongJiangZhuShu1=0 and ZhongJiangZhuShu2=0 and ZhongJiangZhuShu3=0 and ZhongJiangZhuShu4=0 then
		ZhongJiangUser="本期无人中奖"
	else
		ZhongJiangUser="【本期中奖者】"
		ZhongJiangUser=ZhongJiangUser&chr(13)&"一等奖："
		if ZhongJiangZhuShu1>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and HaoMa="&ZhongJiangHaoMa&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang1*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"无  "
		end if
		ZhongJiangUser=ZhongJiangUser&chr(13)&"二等奖："
		if ZhongJiangZhuShu2>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and HaoMa<>"&ZhongJiangHaoMa&" and (HaoMa mod 1000)="&(ZhongJiangHaoMa mod 1000)&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang2*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"无  "
		end if
		ZhongJiangUser=ZhongJiangUser&chr(13)&"三等奖："
		if ZhongJiangZhuShu3>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and (HaoMa mod 1000)<>"&(ZhongJiangHaoMa mod 1000)&" and (HaoMa mod 100)="&(ZhongJiangHaoMa mod 100)&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang3*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"无  "
		end if
		ZhongJiangUser=ZhongJiangUser&chr(13)&"四等奖："
		if ZhongJiangZhuShu4>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and (HaoMa mod 100)<>"&(ZhongJiangHaoMa mod 100)&" and (HaoMa mod 10)="&(ZhongJiangHaoMa mod 10)&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang4*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"无  "
		end if
	end if
	
	dim rs2
	set rs2=server.createobject("adodb.recordset")
	rs2.open "select username from cp_user where Qi="&id&" group by username",conncaipiao,1,3
	rs.open "select * from message where id is null",conn,3,3
	do while not rs2.eof
		rs.addnew
		rs("Sender")="彩票中心"
		rs("incept")=rs2("UserName")
		rs("Title")="第"&ID&"期彩票开奖公告"
		Content="本期中奖号码为"& ZhongJiangHaoMa &"。总计投注"&ZongJiTouZhu&"注，奖池总额"&ZongJinE&"元。其中中奖"& (ZhongJiangZhuShu1+ZhongJiangZhuShu2+ZhongJiangZhuShu3+ZhongJiangZhuShu4) &"注,共得奖金"& JiangJin &"元。"&chr(13)&"本期剩余奖金"&ShengY&"元滚入下期奖池。"
		Content=Content&" "&ZhongJiangUser
		rs("Content")=Content
		rs("issend")=1
		rs.update
		rs2.movenext
	loop
	rs.close
	rs2.close
	set rs2=nothing
	
	'写中奖情况表  我来了 2002.10.16
	rs.open "select * from qingkuang where id is null",conncaipiao,3,3
	rs.addnew
	rs("ZhongJiangHaoMa")=ZhongJiangHaoMa
	rs("ZongJiTouZhu")=ZongJiTouZhu
	rs("ZongJiTouZhujine")=ZongJiTouZhu*JiaGe
	rs("ZongJinE")=ZongJinE
	rs("ZhongJiangZhuShu")=ZhongJiangZhuShu1+ZhongJiangZhuShu2+ZhongJiangZhuShu3+ZhongJiangZhuShu4
	rs("JiangJin")=JiangJin
	rs("ShengY")=ShengY
	rs.update
	rs.close	
end sub

sub savepiao
	dim cp1,cp2,cp3,cp4,xcount,i
	cp1=request("CP1")
	cp2=request("CP2")
	cp3=request("CP3")
	cp4=request("CP4")
	if cp1="" or cp2="" or cp3="" or cp4="" then
		CP_Msg="你至少有一个号码没有购买，请重新选择。"
	end if
	xcount=1
	if (cp1<"0" or cp1>"9") and cp1<>"?" then
		CP_Msg="你至少有一个号码没有购买，请重新选择。"
	else
		if cp1="?" then xcount=xcount*10
	end if
	if (cp2<"0" or cp2>"9") and cp2<>"?" then
		CP_Msg="你至少有一个号码没有购买，请重新选择。"
	else
		if cp2="?" then xcount=xcount*10
	end if
	if (cp3<"0" or cp3>"9") and cp3<>"?" then
		CP_Msg="你至少有一个号码没有购买，请重新选择。"
	else
		if cp3="?" then xcount=xcount*10
	end if
	if (cp4<"0" or cp4>"9") and cp4<>"?" then
		CP_Msg="你至少有一个号码没有购买，请重新选择。"
	else
		if cp4="?" then xcount=xcount*10
	end if
	if usermoney<JiaGe*xcount then
		CP_Msg="对不起，您的社区货币不够购买"&xcount&"注。"
	end if
	if CP_Msg="" then
		rs.open "select * from CP_User where ID is null",conncaipiao,1,3
		for i=0 to 9999
			if (left(right("0000"&i,4),1)=cp1 or cp1="?") and _
				(mid(right("0000"&i,4),2,1)=cp2 or cp2="?") and _
				(mid(right("0000"&i,4),3,1)=cp3 or cp3="?") and _
				(right(right("0000"&i,4),1)=cp4 or cp4="?") then
				rs.addnew
				rs("Username")=Membername
				rs("User")=UserID
				rs("Haoma")=i
				rs("Qi")=CP_Qi
				rs.update
			end if
		next
		rs.close
		conn.execute("update [User] set userwealth=userwealth-"&JiaGe*xcount&" where userid="&userid)
		CP_Msg=cp1&cp2&cp3&cp4&"购买成功，您可以继续选择购买。"
	end if
end sub
%>