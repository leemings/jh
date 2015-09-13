<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select  * from 用户 where 姓名='"&aqjh_name&"'",conn
allvalue=rs("allvalue")
bigdk=allvalue*100
yinliang=rs("银两")
jhdenji=rs("等级")
nowyl=rs("银两")
nowck=rs("存款")
rs.close
rs.open "select * from k where d=false and DateDiff('d',b,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	do while not (rs.bof or rs.eof)
		name=rs("a")
		conn.Execute ("update k set d=true where a='"&name&"'")
		conn.Execute ("delete * from 用户 where 姓名='"& name &"'")
		conn.Execute ("insert into l(b,a,c,d,e) values ('" & name & "',"&sqlstr(sj)&",'高利贷杀手','欠债还钱，没钱要你命','人命')")
		rs.movenext
		name=""
		aqjh_name=""
	loop
end if
%>
<html>
<head>
<title>贷款</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" background="../../bg.gif">
<p align="center"><font size="6" face="隶书">江湖高利贷款</font></p>
<p align="center">&nbsp;</p>
<form method="post" action="dk.asp">
<table width="300" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#E0E0E0">
<tr>
<td>申贷人:</td>
<td><%=aqjh_name%></td>
</tr>
<tr>
<td>现在银两：</td>
<td><%=nowyl%>两</td>
</tr>
<tr>
<td>现在存款:</td>
<td><%=nowck%>两</td>
</tr>
<tr>
<td>最大贷款金额:</td>
<td><%=bigdk%>两</td>
</tr>
<tr>
<td>贷款金额：</td>
<%
rs.close
rs.open "select * from k where a='"&aqjh_name&"' and d=false",conn
if rs.BOF or rs.EOF then
	%>
	<td>
	<%if jhdenji>3 then%>
		<input type="text" name="dk" size="10" maxlength="10">
	<%else%>
		<font color=red>不能操作</font>
	<%end if%>
	</td>
	</tr>
<tr>
<td colspan="2">
<div align="center">
<%if jhdenji>3 then%>
	<input type=submit value="贷款" name="submit">
	<input type="reset" name="reset" value="清空">
<%else%>
	<font color=red>战斗小于3级江湖不贷款给你！</font>
<%end if%>
</div>
</td>
<%else%>
<td>
<%if yinliang>rs("c")*1.5 then%>
<input type="text" name="dk" size="10" maxlength="10" value='<%=rs("c")%>' readonly>
<%else%>
<font color=red>不能操作</font>
<%end if%>
<br>贷款日期：<%=rs("b")%>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<%if yinliang>rs("c")*1.5 then%>
<input type=submit value="还款" name="submit">
<input type="reset" name="reset" value="清空">
<%else%>
<font color=red>你的现金不够还款!</font>
<%end if%>
</div>
</td>
<%end if
rs.close
set rs=nothing
conn.Close
set conn=nothing%>
</tr>
</table>
</form>
<p align="center">本钱庄提供贷款，小人物不贷（战斗等级不到3级不放贷）<br>
贷款期限是一个星期7天。俗话说&quot;欠债还钱，没钱还命&quot;，到期不还者本庄将雇<br>
杀人将其杀之，<font color=red>(将删除江湖ID)</font>望各位大侠相互转告！！！！！！ <br>
(高利贷还钱比例为贷100两，还150两，不是我心黑呀，现在赚钱难呀！) </p>
<p align="center"><font color=red>≮快乐江湖总站≯</font></p>
</body>
</html>
<%
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
%>