<%@ LANGUAGE=VBScript codepage ="936" %>
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("sjjh_name")=""  then response.end
my=Session("sjjh_name")
grade=Session("sjjh_grade")
guess="DBQ="+server.mappath("guess.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set cn=Server.CreateObject("ADODB.CONNECTION")
Set rst=Server.CreateObject("ADODB.RecordSet")
cn.open guess
tempdate=date
rst.open"select * from 中奖号码 where datediff('d','"&tempdate&"',彩票日期)=0",cn,1,1
if rst.bof then
RANDOMIZE
guessno=int((99999-10000+1)*rnd+10000)
'guessno=12345
cn.execute"update 中奖号码 set 中奖号码='"&guessno&"',彩票日期='"&tempdate&"'"
application("guess")=guessno
end if
if application("guess")="" then application("guess")=rst("中奖号码")
sql="SELECT * FROM 中奖号码" 
Set Rst=cn.Execute(sql) 
if not rst.eof then 
	gold=rst("累计金额") 
else 
	gold=0 
end if 
%>
<html>
<head>
<title>六合彩♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="chat/READONLY/Style.css">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
var p,i,x; if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}
function MM_showHideLayers() { //v3.0
var i,p,v,obj,args=MM_showHideLayers.arguments;
for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
obj.visibility=v; }
}
//-->
</script>
</head>
<body bgcolor="#FF9900" background="../chat/bg.jpg">
<div id="Layer1" style="position:absolute; left:60px; top:51px; width:278px; height:203px; z-index:1; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; visibility: hidden"> 
<table width="100%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td>
<p align="center"><font size="+1" color="#0000FF">六合彩说明</font></p>
<p>　<font size="2">　此六合彩每日晚12时开奖，中奖号完全由电脑探制，站长无法操作，本期不中，奖金将移到下一期。每次一人只可购买一个号码，购买号码为5位数字。买一个号码花100两银子，累计金将加1000两。</font></p>
<p align="right"><font size="2">------江湖彩票中心</font></p>
<p align="center">[<a href="#" onClick="MM_showHideLayers('Layer1','','hide')">关闭此窗口</a>]</p>
</td>
</tr>
</table>
</div>
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="32%" height="31">
<tr align="center"> 
<td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">江湖彩票</font></b></td>
</tr>
</table>
<p align="center"> <a href="#" onClick="MM_showHideLayers('Layer1','','show')"> 
<font color="#FFFFFF">购 买 说 明</font></a></p>
<table width="263" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" style=font-size:9pt>
<tr> 
<td width="116">下期奖金己达到</td>
<td width="129"><%=gold%></td>
</tr>
<tr> 
<td width="116">本期中奖号码是</td>
<td width="129"><%=application("guess")%></td>
</tr>
<tr>
<%
sql="SELECT * FROM 购买者 where datediff('d','"&tempdate&"',购买日期)=-1 and 号码="&application("guess")&""
Set Rst=cn.Execute(sql)
if rst.bof then
ren="无人中奖"
else
ren=rst("购买者")
'中奖人加银两
sql="select * from 购买者 where 购买者='"&ren&"' and 领奖=false"
set rst=cn.execute(sql)
if not(rst.bof) then 
sql="update 购买者 set 领奖=true"
set rst=cn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="update 用户 set 银两=银两+"&gold&" where 姓名='"&ren&"'"
set rs=conn.execute(sql)
'加入中奖名
sql="INSERT into 中奖人(中奖者,中奖金额,中奖号码,中奖日期)values('"&ren&"',"&gold&","&application("guess")&",'"&tempdate&"')"
Set Rst=cn.Execute(sql)
'累计金清0
sql="update 中奖号码 set 累计金额=0"
set rst=cn.execute(sql)
end if
end if
sql="select * from 中奖人 where 中奖号码="&application("guess")&" and 中奖者='"&ren&"'"
set rst=cn.execute(sql)
if not(rst.bof) then
ren=rst("中奖者")
jin=rst("中奖金额")
else
ren="无人中奖"
jin="无人中奖"
end if
%>
<tr> 
<td width="116">本期中奖人是</td>
<td width="129"><%=ren%></td>
</tr>
<tr> 
<td width="116">本期中奖金额</td>
<td width="129"><%=jin%></td>
</tr>
</table>
<form method="post" action="buyguess.asp">
<table width="235" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" style=font-size:9pt>
<tr> 
<td width="99">购 买 者</td>
<td width="106"><%=my%></td>
</tr>
<tr> 
<td width="99">您的号码</td>
<td width="106">
<%
sql="select * from 购买者 where 购买者='"&my&"' and datediff('d','"&tempdate&"',购买日期)=0"
set rst=cn.execute(sql)
if rst.bof then
%>
<input type="text" name="guessnumber" size="10" maxlength="10">
</td>
</tr>
<tr> 
<td colspan="2"> 
<div align="center"> 
<input type="submit" name="Submit" value="购买">
　　 
<input type="reset" name="Submit2" value="清空">
</div>
</td>
</tr>
<%
else
myguess=rst("号码")
response.write myguess
end if
%>
</table>
</form>
<p align="center"><font color="#FFFFFF" size="2"> 
<%if ren="无人中奖" then%>
本期无人中奖，奖金将累计到下一期，希望下一期的中奖者就是您! <br>
今天购买，说不准明天中大奖的就是你哟！ 
<%else%>
今日中大奖的用户是<%=ren%>,中奖金额可有<%=jin%>大洋哟!! 
<%end if%>
</font><font color="#FFFFFF"> </font> </p>
<p align=center><font style=font-size:9pt color=#000000>E线江湖</font></p>
</body>
</html>
<%
rst.close
cn.close
set rst=nothing
set cn=nothing
%>