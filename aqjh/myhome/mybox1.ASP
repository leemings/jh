<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title>物品</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#3399CC" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">现有物品一览<font face="幼圆"><a href="javascript:this.location.reload()">刷新</a><br>
  <font color="#FFFF00">玩家可以存储8种物品,每级会员多存4种,房屋每升1级多存4种.</font></font></div>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Response.Expires=0
rs.open "SELECT w1,w2,w3,w4,w5,w6,w7,w8,w9,w10,会员金卡,金币 from [用户] WHERE 姓名='" & aqjh_name & "'",conn,1,1
Response.Write "<table border=""1"" width=""98%"" cellpadding=""0"" cellspacing=""0"" height=""5"" bordercolordark=""#FFFFFF"" bordercolorlight=""#00215a"" align=""center"">"
Response.Write "<tr align=""center"">"
tt=0
for t=1 to 9
if not(isnull(rs("w" & t))) then
	data1=split(rs("w" &t),";")
	data2=UBound(data1)
	for y=0 to data2-1
	tt=tt+1
	data3=split(data1(y),"|")
	Response.Write "<td width=""84"" height=""3""><A href=""javascript:show('"& data3(0) &"','"& data3(1) &"','w"& t &"');""><font color=#ffffff>"& data3(0) &"</font></a> "& data3(1) &"个</td>"
	if tt/4=int(tt/4) then Response.Write "</tr><tr align=""center"">"
	erase data3
	next 
	erase data1
end if
next
tt=tt+1
Response.Write "<td width=""84"" height=""3""><A href=""javascript:show('会员金卡','"& rs("会员金卡") &"','会员金卡');""><font color=#ffffff>金卡</font></a> "& rs("会员金卡") &"元</td>"
if tt/4=int(tt/4) then Response.Write "</tr><tr align=""center"">"
tt=tt+1
Response.Write "<td width=""84"" height=""3""><A href=""javascript:show('金币','"& rs("金币") &"','金币');""><font color=#ffffff>金币</font></a> "& rs("金币") &"个</td>"
if tt/4=int(tt/4) then Response.Write "</tr><tr align=""center"">"
Response.Write "</table>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<SCRIPT language=JavaScript>
function show(name,mysl,lb){
parent.cw.sl=mysl;
parent.cw.document.form.name.value=name;
parent.cw.document.form.wpsl.value=mysl;
parent.cw.document.form.lb.value=lb;
parent.cw.document.form.Submit.value="放入";
}
</SCRIPT>
</body>
</html>
