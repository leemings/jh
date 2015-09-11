<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
wpsl=mywpsl(rs("w6"),"冰水")
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>;parent.show.fm1.innerHTML="&chr(34)&"<div align='center'>"&wpsl&"<img src='images/tong4.gif'></div>"&chr(34)&";parent.show.fm2.innerHTML="&chr(34)&"<div align='center'>0<img src='images/tong4.gif'></div>"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"<div align='center'><img src='images/tong0.gif'></div>"&chr(34)&"</script>"
%>
<HTML><HEAD><TITLE>爱情江湖 永不放弃制作</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<SCRIPT src="images/icenew.js"></SCRIPT>
<META content="MSHTML 5.00.2920.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 leftMargin=0 topMargin=0>
<FORM action=jl.asp name=shiform target=cz>
</FORM>
<DIV id=Layer1
style="HEIGHT: 14px; LEFT: 12px; POSITION: absolute; TOP: 8px; WIDTH: 64px; Z-INDEX: 1"><IMG
height=14 src="images/zi.gif" width=64></DIV>
<DIV id=Layer2
style="HEIGHT: 23px; LEFT: 472px; POSITION: absolute; TOP: 12px; WIDTH: 106px; Z-INDEX: 2"><A
href="#"
onmousedown=go()><IMG alt=采冰 border=0 height=21 src="images/caibin.gif"
width=20></A> &nbsp; <A
href="#"
onmousedown=to()><IMG align=middle alt=化冰 border=0 height=17
src="images/huabin.gif" width=17></A> &nbsp; <A
href="#"
onclick=><IMG alt=出口 border=0 height=27
src="images/exit.gif" width=20></A> &nbsp; </DIV>
<IMG id=yan src="images/Yan.gif"
style="LEFT: 113px; POSITION: absolute; TOP: 189px; VISIBILITY: hidden; Z-INDEX: 3">
<IMG id=huo src="images/D5106.gif"
style="LEFT: 147px; POSITION: absolute; TOP: 312px; VISIBILITY: hidden; Z-INDEX: 4">
<DIV style="LEFT: -1px; POSITION: absolute; TOP: 0px; Z-INDEX: -1; width: 596px; height: 436px"><IMG border=0
height=436 src="images/back.jpg" width=600></DIV>
<IMG id=person
src="images/go0.gif"
style="LEFT: 200px; POSITION: absolute; TOP: 315px; Z-INDEX: 33">
<div id=Layer1 style="HEIGHT: 14px; LEFT: 12px; POSITION: absolute; TOP: 402px; WIDTH: 588px; Z-INDEX: 1"><font size="-1"><b><font color="#0000FF">点击锤子开始采冰，点绿色方块开始化冰取得冰水！</font></b></font></div>
</BODY>
</HTML>