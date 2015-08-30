<%@ LANGUAGE=VBScript codepage ="936" %>
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
sjjh_name=sjjh_name
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
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
sj1=n & "-" & y & "-" & r
sj2=s & ":" & f & ":" & m
sj3=sj1 & " " & sj2
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
Application("song")="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("eline__SONG.mdb")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("song")
Set Rs=conn.Execute("select  * FROM song")
if chatbgcolor="" then chatbgcolor="008888"%>

<html>
<head>
<title>点歌</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<script language="javaScript">
function send(){sendurl="songplay.asp?name=" + document.forms[0].song.options[document.forms[0].song.selectedIndex].value;window.location.href=sendurl;}
</script>
<link rel="stylesheet" href="readonly/style.css">
<style type="text/css">
td           { font-family: 宋体; font-size: 12px }
body         { font-family: 宋体; font-size: 12px;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
</style>
</head>
<body bgcolor="#006699" text="#FFFFFF" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#CCCCFF" size="2"><strong>我要点歌</strong></font></div></td>
      </tr>
    </table>
<table width="140" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#000000"
bordercolordark="#FFFFFF" bgcolor="#006699">
  <form method="post" action="songplay.asp?" name="" target="ps">
<tr>
<td>

          <p align="center"><a href="javascript:history.go(0)" title="显示最新的曲目列表"><font color="#FFFFFF">刷 新</font></a></p>

        <p> <font color="#FFFF00">选择曲目</font><br>
          <select name="song" size="8" style="font-size:9pt">
            <%do while not rs.bof and not rs.eof %>
            <option value="<%=rs("name")%>"><%=rs("name")%></option>
<%
rs.movenext
loop
rs.close
conn.close
set rs=nothing
set conn=nothing
%>





erase mid    
</select>
</p>
        <p><font color="#FFFF00">播放方式</font><br>
<input type="radio" name="loopok" value="1" checked>
          听一遍(稍等一会)<br>    
<input type="radio" name="loopok" value="infinite">
          听不厌(稍等一会)<br>    
<input type="radio" name="loopok" value="ddj">
给大家(需03金币)<br>    
<input type="radio" name="loopok" value="dhy">
          给好友(需10银币)<br>    
<br>
          <font color="#ffff00">祝福话语</font><br>
<!--webbot bot="Validation" B-Value-Required="TRUE" I-Maximum-Length="50" --><input type="text" name="zhufu" size="17" style="BACKGROUND-COLOR: #666666; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #00ff00; FONT-SIZE: 9pt" value="写上你要说的话" maxlength="50"><br>
          <font color="#ffff00">朋友名字</font><br>
    <select name="to1" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px">
      <%useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
for i=0 to x
if show(i)<>sjjh_name and show(i)<>"回首当年" and show(i)<>"无非" and show(i)<>"毒药" Then
%>
      <option value="<%=show(i)%>" selected><%=show(i)%></option>
<%end if
next%>
<%if x=0 then%>
      <option value="就你自己一个人" selected>就你自己一个人</option>
<%end if%>
    </select>


 　</p>
<table border="0" cellpadding="4">
<tr>
<td>
<input type="submit" name="play" value="播放" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '宋体'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
</td>
<td align="right">
<input type="submit" name="st" value="停止" style="BACKGROUND-COLOR: #336699; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '宋体'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
<Script Language=Javascript>
parent.m.location.href='about:blank';
</Script>
</body>
</html>