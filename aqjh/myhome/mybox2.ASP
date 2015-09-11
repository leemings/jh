<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
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
<title>储物箱</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#339999" text="#FFFFFF" leftmargin="0" topmargin="8">
<div align="center">现有储物箱一览<font face="幼圆"><a href="javascript:this.location.reload()">刷新</a><br>
  <font color="#FFFF00">会员到期后,超过种类的物品会自动加到身上.</font></font></div>
<table border="1" width="98%" cellpadding="0" cellspacing="0" height="5" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
<tr align="center"><form action="mybox3.asp" method="post" name="form" target="cz" id="form" onsubmit='return(check());'>
<td colspan=4>物品名<input name="name" readonly type="text" id="name" size="8">
数量<input name="wpsl" type="text" id="wpsl" size="5"> 
第二密码 
<input name="pass" type="password" id="pass" size="8" maxlength="50">
<input name="lb" type="hidden" id="lb">
<input type="submit" name="Submit" value="放入">
</td></form>
</tr>
<%
Response.Write "<tr align=""center"">"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.Open "select * from house  where "& session("myroom") &"='"& session("aqjh_name") &"'", conn, 1, 1
if rs.Eof and rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
   	Response.Write "<script language=JavaScript>{alert('提示:没有购买房子是不能使用储物箱的！');}</script>"
	Response.End
end if
roomjb=rs("h_级别")-1
rs.close
rs.open "SELECT * from mybox WHERE mb_账号='" & session("aqjh_name") & "'",conn,1,1
records=rs.recordcount
rs.close
rs.open "SELECT 会员等级 from 用户 WHERE 姓名='" & session("aqjh_name") & "'",conn,1,1
hydj=rs("会员等级")
nowsl=8+hydj*4+roomjb*4
rs.close
if records>nowsl then
	Set rs1=Server.CreateObject("ADODB.RecordSet")
	rs1.open "SELECT w1,w2,w3,w4,w5,w6,w7,w8,w9,会员金卡,金币 from 用户 WHERE 姓名='" & session("aqjh_name") & "'",conn,1,3
	rs.open "SELECT * from mybox WHERE mb_账号='" & session("aqjh_name") & "'",conn,1,3
	xx=0
	do while xx<(records-nowsl)
		xx=xx+1
		if trim(rs("mb_物品名"))="会员金卡" or trim(rs("mb_物品名"))="金币" then
			rs1(trim(rs("mb_类别")))=rs1(trim(rs("mb_物品名")))+rs("mb_数量")
		else
			zstemp=add(rs1(trim(rs("mb_类别"))),trim(rs("mb_物品名")),clng(rs("mb_数量")))
			if zstemp="1" then
				rs.Close
				Set rs = Nothing
				conn.Close
				Set conn = Nothing
				Response.Write "<script Language=Javascript>alert('提示：你的物品超过25种不能保存？');</script>"
				response.end
			end if
			rs1(trim(rs("mb_类别")))=zstemp
		end if
		rs1.update
		rs.delete
		rs.update
		rs.movenext
	loop
	rs1.close
	rs.close
	set rs1=nothing
	set rs=nothing
	conn.close
	Response.Write "<script language=JavaScript>{alert('提示:物品种类超过限制,超过部份放到你的物品包中！');parent.cw.location.reload();parent.wp.location.reload();}</script>"
	response.end
end if
rs.open "SELECT * from mybox WHERE mb_账号='" & session("aqjh_name") & "'",conn,1,1
tt=0
do while not rs.bof and not rs.eof
	tt=tt+1
	Response.Write "<td width=""84"" height=""3""><A href=""javascript:show('"& trim(rs("mb_物品名")) &"','"& rs("mb_数量") &"','"& trim(rs("mb_类别")) &"');""><font color=#ffffff>"
	if trim(rs("mb_物品名"))="会员金卡" then
		Response.Write "金卡"
	else
		Response.Write trim(rs("mb_物品名"))
	end if
	Response.Write "</font></a> "& rs("mb_数量") &"个</td>"
	if tt/4=int(tt/4) then Response.Write "</tr><tr align=""center"">"
rs.movenext
loop
Response.Write "</table>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<SCRIPT language=JavaScript>
var sl=0;
function show(name,mysl,lb){
parent.cw.sl=mysl;
parent.cw.document.form.name.value=name;
parent.cw.document.form.wpsl.value=mysl;
parent.cw.document.form.lb.value=lb;
parent.cw.document.form.Submit.value="取出";
}
function check(){
var pattern = /^([0-9])+$/;
var name=document.form.name.value;
var wpsl=document.form.wpsl.value;
var lb=document.form.lb.value;
var cz=document.form.Submit.value;
var pass=document.form.pass.value;
if (cz=="取出" && pass==""){alert("提示：取出物品时请输入第二密码!");return false;}
if (name==""){alert("提示：名字不能为空！");return false;}
if(pattern.test(wpsl)!=true){alert("提示：数量请使用数字！");return false;}
if(wpsl==0){alert("提示：物品数量不能为0！");return false;}
if(parseFloat(sl)!=0 && parseFloat(wpsl)>parseFloat(sl)){alert("提示：你的物品只有"+sl+"个！");return false;}
return true;
}
</SCRIPT>
</body>
</html>




