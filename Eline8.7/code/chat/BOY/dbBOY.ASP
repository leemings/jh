<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

rs.open "select 配偶 from 用户 where 姓名='"&sjjh_name&"'",conn,3,3
if isnull(rs("配偶")) or rs("配偶")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对象都没有，想自己一个人生呀！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select 结婚记念日 from 用户 where 姓名='"&sjjh_name&"'",conn,3,3
if DateDiff("d",rs("结婚记念日"),date())<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的爱情还不够稳定，等2天后吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select boy,boysex from 用户 where 姓名='"&sjjh_name&"'",conn,3,3
if isnull(rs("boy")) or rs("boy")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还在恋爱中，还没有生小孩！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
mypic=rs("boysex")

zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩数据出错，请重新生育！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if clng(zt(3))<=0 then
	conn.execute "update 用户 set boy='' where  姓名='"&sjjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：因为你不好好照顾小孩,小孩病死了！！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if


cwsm=clng(DateDiff("d",zt(7),now()))
if cwsm>0 then
	temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-cwsm*15)&"|"&zt(4)&"|"&zt(5)&"|"&zt(6)&"|"&zt(7)
	conn.execute "update 用户 set boy='"&temp&"' where  姓名='"&sjjh_name&"'"
		if (clng(zt(3))-cwsm*15)<=0 then
			conn.execute "update 用户 set boy='' where  姓名='"&sjjh_name&"'"
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：因为你不好好照顾小孩,小孩病死了！！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
			response.end
		end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
zgg=(4+int(DateDiff("d",zt(2),now())/10))
cssj=clng(DateDiff("d",zt(2),now()))
gs=clng(zt(5))
if gs<(80+cssj*5) then
	czzt="叛逆"
elseif gs>((100+cssj*5+50)) then
	czzt="顺从"
else
	czzt="普通"
end if
%>
<html>
<head>
<title>我的小孩</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
body{
CURSOR: url('boy.cur');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#006699" bgproperties="fixed" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td width="100%" height="28"> <div align="center"><font color="#ccccff" size="2"><b>我的小孩</b></font></div></td>
  </tr>
</table>
<table background="../../bg2.gif" border="1" cellspacing="0" width="140" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr> 
    <td>照片:<img src="<%=mypic%>" width=40 height=40><br><br>
      名字:<%=zt(0)%><br><br>
      职业:正在开发<br><br>
      体力:<%=(clng(zt(3))-cwsm*15)%><br><br>
      攻击:<%=zt(4)%><br><br>
      感受:<%=zt(5)%><br><br>
      状态:<%=czzt%><br><br>
      照顾:<%=zg%>/<font color=red><b><%=zgg%></b></font><br><br>
      生日:<%=Year(zt(2))%>-<%=Month(zt(2))%>-<%=day(zt(2))%><br><br>
      出生:<%=DateDiff("d",zt(2),now())%>天<br><br>
    </td>
  </tr>
  <tr>
    <td><div align="center"><a href="javascript:s('/夺宝小孩攻击$ <%=zt(0)%>')" ><font color="#0066FF">夺宝攻击</font></a></div></td>
  </tr>
</table>
</body>
</html>