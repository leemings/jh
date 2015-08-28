<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
diaotime=100
nowtime=now()
lastlogintime=dateadd("s",diaotime,nowtime)
lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
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
	sj=n & "-" & y & "-" & r 
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="Select 积分,银两,形态 from 用户 where 姓名='"&username&"'"  
set rs=conn.Execute(sql)
allvalue=rs("积分")
yinliang=rs("银两")
xingtai=rs("形态")
if xingtai="鬼魂" then 
set rs=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你是一只鬼魂，还要来贷款？');window.close();</script>"
Response.end
end if
bigdk=int(yinliang*0.05-allvalue*4)
if bigdk<0 then bigdk=int(yinliang*0.05)
%>
<!--#include file="data.asp"--><%
	sql="select * from 贷款 where 还贷记录=false and DateDiff('d',贷款日期,#" & sj & "#)>7"
	Set Rs=connt.Execute(sql)
	if not(rs.BOF or rs.EOF) then 
	do while not (rs.bof or rs.eof)
	name=rs("贷款人")
	connt.Execute ("update 贷款 set 还贷记录=true where 贷款人='"&name&"'")
	conn.Execute ("update 用户 set 体力=体力-100000,内力=内力-100000,形态='鬼魂',最后登录时间="&lastlogintimetype&" where 姓名='"&name&"'")
	conn.Execute ("insert into 英烈堂(死者,时间,凶手,死因) values ('" & name & "',now(),'剑侠杀手','贷款过期不还')")
	session.Abandon
	name=""
	rs.movenext
	loop
	end if
%>
<html>
<head>
<title>贷款</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../Style.css">
</head>

<body background='../../chatroom/bg1.gif'>
<p align="center"><font face="隶书" size="4"><b>江湖高利贷</b></font></p>
<form method="post" action="dk.asp">
  <table width="300" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#E0E0E0">
    <tr> 
      <td>申贷人:</td>
      <td><%=username%></td>
    </tr>
    <tr>
      <td>现在银两：</td>
      <td><%=yinliang%>两</td>
    </tr>
    <tr> 
      <td>最大贷款金额:</td>
      <td><%=bigdk%>两</td>
    </tr>
    <tr> 
      <td>贷款金额：</td>
      <%
	sql="select * from 贷款 where 贷款人='"&username&"' and 还贷记录=false"
	set rs=connt.execute(sql)
	if rs.BOF or rs.EOF then
     %> 
      <td> 
        <input type="text" name="dk" size="10" maxlength="10">
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center"> 
          <input type=submit value="贷款" name="submit">
          <input type="reset" name="reset" value="清空">
        </div>
      </td>
      <%else%> 
      <td> 
        <input type="text" name="dk" size="10" maxlength="10" value='<%=rs("贷款金额")%>' readonly>
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center"> 
          <input type=submit value="还款" name="submit">
          <input type="reset" name="reset" value="清空">
        </div>
      </td>
      <%end if
	  conn.Close 
	  set conn=nothing
	  connt.Close 
	  set connt=nothing%> </tr>
  </table>
</form>
<p align="center">本钱庄提供贷款，小人物不贷（等级不到20级不放贷），贷款期限是一个星期。<br>
  俗话说&quot;欠债还钱，没钱还命&quot;，到期不还者本庄将雇<br>
杀手将其杀之，<font color="red"><b>(将成为江湖鬼魂,每次登陆都会减少5％属性)</b></font>望各位大侠相互转告！！！！！！<br>
高利贷还钱比例为：会员贷100两，还120两，<br>非会员贷100两还150两不是我心黑呀，现在赚钱难呀！ </p>
</body>
</html>
