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
wpsl=mywpsl(rs("w6"),"木头")
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>;parent.show.fm1.innerHTML="&chr(34)&wpsl&chr(34)&";parent.show.fm2.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"0"&chr(34)&"</script>"
%>
<HTML><HEAD><TITLE>高山小区筏木</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<SCRIPT src="image/stone.js"></SCRIPT>
<script Language=Javascript>

</SCRIPT>
<STYLE>.csstext1 {
COLOR: #ffffff; FILTER: glow(Color=red,Strength=#1)'}; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; LINE-HEIGHT: normal
}</STYLE>

<STYLE>.csstext2 {
COLOR: #ffffff; FILTER: glow(Color=gold,Strength=#1)'}; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; LINE-HEIGHT: normal
}
</STYLE>
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=#c0c0c0 leftMargin=0  topMargin=0 ;>
<TABLE background=image/index.jpg border=0 cellPadding=0 cellSpacing=0
height=368 width=500>
<TBODY>
<TR>
<TD vAlign=bottom width="100%"><IMG border=0 height=45 id=caikuang
src="image/d.gif"
style="LEFT: 300px; POSITION: absolute; TOP: 91px; VISIBILITY: hidden"
width=33>
<TABLE border=0 borderColorLight=#735639 cellPadding=0 cellSpacing=0
width="100%">
<TBODY>
<TR>
<TD align=right width="101%">
<TABLE background=image/back2.gif bgColor=#000000 border=0
cellPadding=0 cellSpacing=0 height=29 width="100%">
<TBODY>
<TR>
<TD align=middle width="8%">　</TD>
<TD align=middle class=csstext1 id=kstext width="8%">　</TD>
<TD align=middle width="8%">　</TD>
<TD align=middle class=csstext1 id=tktext width="8%">　</TD>
<TD align=middle width="8%">　</TD>
<TD align=middle class=csstext1 id=ksdj width="8%">　</TD>
<TD align=middle width="8%">　</TD>
<TD align=middle class=csstext1 id=tkdj width="8%">　</TD>
<TD align=middle width="9%">　</TD>
<TD align=middle class=csstext1 id=jingyan width="9%">　</TD>
<TD align=middle width="9%">　</TD>
<TD align=middle class=csstext1 id=jineng
width="9%">　</TD>
</TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<P><IMG border=0 height=45 id=liantie src="image/c.gif"
style="LEFT: 298px; POSITION: absolute; TOP: 82px; VISIBILITY: hidden"><SPAN
id=dostone style="LEFT: 381px; POSITION: absolute; TOP: 130px"><A
href="#"
onclick="setTimeout('gocai()',0)"> <IMG alt=筏木 border=0 class=csstext2 height=29
src="image/caikuang.gif" style="POSITION: absolute" width=30></A></SPAN>
<SPAN id=dosteel style="LEFT: 35px; POSITION: absolute; TOP: 180px"></SPAN></P>
<DIV id=sjdiv style="border:1px none #000000; BACKGROUND-IMAGE: url('/tribe/hut/stone/back3.gif'); HEIGHT: 100px; LEFT: 250px; POSITION: absolute; TOP: 160px; VISIBILITY: hidden; WIDTH: 100px; zindex: 100; layer-background-image: url(/tribe/hut/stone/back3.gif)"><BR>
<FORM action=jl.asp name=shiform target=ig>
</FORM>
<FORM name=form1>
<INPUT CHECKED id=actbutton1 name=sjbutton type=radio value=3>
<LABEL for=actbutton1><SPAN style="LEFT: 20px; POSITION: absolute; TOP: 19px">
<IMG border=0 height=19 src="image/cd.gif" width=76></SPAN></LABEL><BR>
<INPUT id=actbutton2 name=sjbutton type=radio value=4>
<LABEL for=actbutton2><SPAN style="LEFT: 20px; POSITION: absolute; TOP: 42px">
<IMG border=0 height=19 src="image/ld.gif" width=76></SPAN></LABEL>
</FORM>
<A href="#" onclick=getJnd();>
<P align=center><IMG border=0 height=15 src="image/sj.gif"
style="fliter: fire" width=53></P></A></DIV></BODY></HTML>