<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("aqjh_usermdb")
Set rst= Server.CreateObject("ADODB.Recordset")
Set rs=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT w6,w8,w9,w10,w3 FROM 用户 WHERE 姓名='" & aqjh_name& "'"
rst.open sqlstr,conn,1,2
if IsNull(rst("w6")) and IsNull(rst("w8")) and IsNull(rst("w9")) and IsNull(rst("w10")) and IsNull(rst("w3")) then
	rst.close
	set rst=nothing
	conn.close
	set conn=nothing
	msg=""
else
d6=rst("w6")
d8=rst("w8")
d9=rst("w9")
d10=rst("w10")
d3=rst("w3")
msg=""
if not IsNull(d6) then
	call show(d6,"物品",msg)
end if
if not IsNull(d8) then
	call show(d8,"配药",msg)
end if
if not IsNull(d9) then
	call show(d9,"魔宝",msg)
end if
if not IsNull(d10) then
	call show(d10,"法器",msg)
end if
if not IsNull(d3) then
	call show(d3,"装备",msg)
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="你没有东西可供寄售！"
%>
<head><title>自由集市</title><LINK href="../../css.css" rel=stylesheet><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background="../../bg.gif"  oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='购买集市中流通的物品';return true;" onmouseout="window.status='';return true">购买</a></td>
<td bgcolor=#ffff00><A href="sale.asp" onmouseover="window.strtus='标价出售自己拥有的物品';return true;" onmouseout="window.status='';return true">寄售</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("aqjh_chatroomname")%>自由集市</b></font></p>
<p align=center>注：商店可以买到的东西禁止出售</p>
<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=ffffff align='center'><tr bgcolor=ffff00 align=center><td width=40>类型</td><td width=140>身上物品</td><td>物品数量</td><td>单价</td><td width=36>寄售</td></TR>
<%=msg%></table>
<p align=center>
<input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p>
</body>
<%
sub show(data,lx,msg)
data1=split(data,";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
	rs.open "SELECT * FROM b WHERE a='" & data3(0)& "'",conn
	if not rs.bof and not rs.eof then
msg=msg&"<tr><td>"&lx&"</td><td>"&data3(0)&"</td><td>"&data3(1)&"</td><td>"&rs("h")&"</td><td><a href='vend.asp?lx="&lx&"&name="&data3(0)&"&sl="&data3(1)&"&danjia="&rs("h")&"'>寄售</a></td></tr>"
  else
msg=msg&"<tr><td>"&lx&"</td><td>"&data3(0)&"</td><td>"&data3(1)&"</td><td>10000</td><td><a href='vend.asp?lx="&lx&"&name="&data3(0)&"&sl="&data3(1)&"&danjia=10000'>寄售</a></td></tr>"
  end if
  rs.close
erase data3
next
erase data1
end sub
%>