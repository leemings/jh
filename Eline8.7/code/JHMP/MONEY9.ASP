<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
id=LCase(trim(request.querystring("id")))
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 姓名='"&sjjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("会员等级")=1  then
Response.Write "<script language=JavaScript>{alert('你是等级会员,不应该来这种地方的!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("会员等级")=2  then
Response.Write "<script language=JavaScript>{alert('你你是等级会员,不应该来这种地方的!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("会员等级")=3  then
Response.Write "<script language=JavaScript>{alert('你是等级会员,不应该来这种地方的!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("会员等级")=4  then
Response.Write "<script language=JavaScript>{alert('你是等级会员,不应该来这种地方的!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("会员等级")=5  then
Response.Write "<script language=JavaScript>{alert('你是等级会员,不应该来这种地方的!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if Session("sjjh_grade")>=5 then Response.Write "<script language=JavaScript>{alert('掌门和官府是不应该来这种地方的!');location.href = 'javascript:history.back()';}</script>"
pai=rs("门派")
rs.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
dj=rs("等级")
if dj<20 then
Response.Write "<script language=JavaScript>{alert('你战斗等级不够20级不能进来领取金币!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if


Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'以上是对联盟网吧判断
wbname=""
wbpf=0
ip=Request.ServerVariables("REMOTE_ADDR")
rs.open "SELECT barname FROM bar WHERE ip='"&ip&"'",conn
if Not(rs.Eof and rs.Bof) then
wbname=rs("barname")
wbpf=1
end if
rs.close
tmprs=conn.execute("Select count(*) As 数量 from 用户 where allvalue>100 and 介绍人='"&sjjh_name&"'")
jsren=tmprs("数量")
rs.open "Select * from 用户 where 姓名='"&sjjh_name&"'",conn
mdate=rs("领钱")
'if CDate(mdate)=now() then
'if day(mdate)>=day(now()) and month(mdate)=month(now()) and year(mdate)=year(now()) then 
if  DateDiff("d",rs("领钱"),now())=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&sjjh_name &"]今天你来领过钱的！忘记了？');window.close();</script>"
	response.end
end if
yin=rs("银两")+rs("存款")
if yin>(rs("会员等级")+1)*500000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&sjjh_name &"]你现在有这么多的银两你还想要呀！！');window.close();</script>"
	response.end
end if
dj=rs("等级")
hy=rs("会员等级")
if trim(rs("身份"))="掌门" then
	'如是掌门统计人数
	tmprs=conn.execute("Select count(*) As 数量 from 用户 where 门派='"&rs("门派")&"'")
	renshu=tmprs("数量")
	'算钱数
	money=cint(renshu)*100000+jsren*500000+dj*50000+wbpf*100000
	conn.execute("update 用户 set 银两=银两+"&money&",领钱=now() where 姓名='"&sjjh_name&"'")
else
	'个人按会员乘10万
	money=jsren*500000+dj*500000+wbpf*3000000+1000000
	conn.execute("update 用户 set 银两=银两+"&money&",领钱=now() where 姓名='"&sjjh_name&"'")
end if
rs.close
rs.open "Select * from 用户 where 姓名='"&sjjh_name&"'",conn
conn.execute "update 用户 set 金币=金币+1 where 姓名='"&sjjh_name&"'"
%>
<html>
<head>
<title>工资领取处</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
BODY {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
SELECT {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#CCCCCC" text="#000000" leftmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="center">
<p><font color="#000000"><b><font size="+3">工资领取处</font></b></font> <br>
<br>
<font size=+1> <b><%=sjjh_name%></b></font>你是<font color="#0000FF">[<%=rs("门派")%>]</font>
的<font color="#FF0000"> [<%=rs("身份")%>] </font>拉人：<font color="#FF0000"><b><%=jsren%>个</b></font>      
战斗等级:<font color="#FF0000"><%=rs("等级")%>级</font>      
<%if trim(rs("身份"))="掌门" then%>      
你有弟子:<font color="#FF0000"><b><%=renshu%>个</b></font> <%end if%><br>多多努力，多多拉人！      
<br>
<%if wbpf<>0 then%>
 <font color=blue>您在联盟网吧：[<%=wbname%>]上网，多领300万！</font><br>    
<%end if%>
今天你从江湖领取到了<b><font color="#FF0000"><%=money%>两</font></b>，小心保存，不要乱花！      
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
<br>
</p>
<p align="center"><br>
<font color="#FF0000" size="+1"><b>工资计算方法</b></font><br>
<font color="#0000FF">弟子：</font>介绍人数x50万+战斗等级x50万+100万</p>
<p align="center"><b><font color="#FF00FF">战斗等级20级以上另外送金币<img border="0" src="Jb.GIF">1个作为奖励。</font></b></p>
<p align="center">&nbsp;</p>
<p align="center">
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</div>
</body>
</html>
