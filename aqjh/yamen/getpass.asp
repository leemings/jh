<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
randomize timer
regjm=int(rnd*9998)+1
%>
<HTML><HEAD><TITLE>取回密码</TITLE>
<LINK href="pic/lyy.css" rel=stylesheet></HEAD>
<BODY background=pic/BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="getpassok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=380 align=center border=0>
<TBODY><TR>
<TD width="3%" height=1><IMG src="pic/T1.gif" border=0></TD>
<TD width="94%" background=pic/M1.gif height=1>　</TD>
<TD width="12%" height=1><IMG src="pic/T2.gif" border=0></TD></TR>
  <TR><TD width="3%" background=pic/M2.gif rowSpan=3></TD>
    <TD width="94%" height=23><P align=center><SPAN lang=en>&copy;
<B>请大侠仔细填写以下信息------取回密码</P></SPAN></TD>
    <TD width="12%" background=pic/M2.gif height=1 rowSpan=3></TD></TR>
  <TR>
    <TD width="94%" height=16>
      <HR color=#682420 SIZE=1>
    </TD></TR>
  <TR><TD width="94%">
<!=>
<table width="254" align="center">
          <tr>
            <td>
<div align="center"><br>侠客ID号：
                <input type="text" name="userid" size="10" maxlength="10" class=input><br>
                侠客姓名：
                <input type="text" name="name" size="10" maxlength="10" class=input>
                <br>
                 第二密码：
                <input type="password" name="pass" size="10" maxlength="10" class=input>
                <br>
                输入认证：
                <input type=text name=regjm1 size=5 maxlength="5" value='<%=regjm%>' class=input>
                <br>
                认证号：<font color="#FF0000"><%=regjm%></font> <br>
<input type=hidden name=regjm value="<%=regjm%>">
                <p> 
</div></td></tr>
<tr><td align="center">
<input type="submit" name="submit" value="确定" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8">
<input type="button" value="关闭" onclick="window.close()"
name="button" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8"></td>
</tr><TR>
          <TD align=middle><FONT 
            color=#0000ff>&copy; 版权所有 2004-2005 </FONT><A 
            href="http://www.7758530.com/" target=_blank><FONT 
            color=#0000ff>爱情江湖网</FONT></A> </TD></TR>
</table>
<!=></TD></TR></form>
<TR>
    <TD width="3%" height=6><IMG src="pic/T3.gif" border=0></TD>
    <TD width="94%" background=pic/M3.gif height=6>　</TD>
    <TD width="12%" height=6><IMG src="pic/T4.gif" 
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>