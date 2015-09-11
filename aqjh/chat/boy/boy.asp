<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 配偶,性别 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
mysex=rs("性别")
if isnull(rs("配偶")) or rs("配偶")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对象都没有，想自己一个人生呀！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select 结婚记念日 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if DateDiff("d",rs("结婚记念日"),date())<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的爱情还不够稳定，等2天后吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select boy,boysex from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
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
	conn.execute "update 用户 set boy='' where  姓名='"&aqjh_name&"'"
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
	conn.execute "update 用户 set boy='"&temp&"' where  姓名='"&aqjh_name&"'"
		if (clng(zt(3))-cwsm*15)<=0 then
			conn.execute "update 用户 set boy='' where  姓名='"&aqjh_name&"'"
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
<script Language="Javascript">
function del(){if(confirm("您真的要丢弃你的小孩吗？")){window.location.href="delboy.asp";return true;}}
</script>
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
<body oncontextmenu=self.event.returnValue=false bgproperties="fixed" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" background="../<%=chatimage%>">
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
      性别:<%=zt(1)%><br><br>
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
    <td><div align="center"><a href="yang.asp" title="给小孩吃点米饭吧，增加体力。"><font color="#ffffff">养</font></a> <a href="dlboy.asp" title="锻炼小孩增加攻击!" target="d"><font color="ffffff">锻</font></a> 
        <a href="pbboy.asp" title="陪伴小孩，增加感受!" target="d"><font color="ffffff">陪</font></a> <a href="gmboy.asp" title="取个好名字吧!" ><font color="ffffff">改</font></a> <a href="javascript:s('/小孩$ <%=zt(0)%>')" ><font color="ffffff">攻</font></a> <a href="javascript:s('/小孩护主$ <%=zt(0)%>')" title="小孩护主">暴</a></div></td>
  </tr>
<%if mysex="女" then%>
<tr><td align=center><a href=# onclick="javascript:del()">丢弃小孩</a></td></tr>
<%end if%>
</table>
</body></html>