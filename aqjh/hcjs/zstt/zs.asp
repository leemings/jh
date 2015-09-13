<%@ LANGUAGE=VBScript codepage ="936" %>
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
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>脱胎换骨，得道飞升</TITLE>
<BGSOUND src="image/baihe.mid" loop=-1>
<link rel="stylesheet" type="text/css">
<style>
<!--
.6     { background-color: #E1E9FF; vertical-align: 1 }
-->
</style>
</HEAD>
<BODY vLink=#e1e9ff aLink=#e1e9ff link=#e1e9ff bgColor=#78869f background="image/wonderofitallbkgtile2a.jpg" style="font-size: 8pt; color: #FFFFFF">
<DIV align=center>
<CENTER>
<TABLE cellSpacing=20 borderColorDark=#3d4b61 cellPadding=0 width="88%" borderColorLight=#8f9fba background="image/wonderofitallstrip.jpg" border=1>
<TBODY> 
<TR>
<TD width="100%" align=center>
<TABLE cellSpacing=10 borderColorDark=#3d4b61 cellPadding=0 width="100%" borderColorLight=#8f9fba background="image/wonderofitalltilelight.jpg" border=1>
 <TBODY> 
 <TR>
 <TD width="100%" height="411" align=center> 
<TABLE cellPadding=0 width="100%" background="image/wonderofitallstrip.jpg" border=0>
 <TBODY>
 <TR>
 <TD width="100%" align=center>
<TABLE cellPadding=0 width="100%" background="image/wonderofitalltile.jpg" border=0>
<TBODY>
 <TR align=center>
<TD width="100%" height="372" valign="top"> 
<table width="594" border="0" cellspacing="0" cellpadding="0" height="25">
<tr> 
<td width="235" height="31" valign="middle" align="center"><br>
<br>
                                          <font color="FFFFFF" zize=+3"><b>¤脱胎换骨，得道飞升¤</font> 
                                        </td>
<td width="359" height="235" rowspan="2"> <img height=274 src="image/right.jpg" width=360 border=0></td>
</tr>
<tr>
<td width="235" height="215" valign="top">
<br><br>
                                          <font size="2"><font color="#FFFF00"><b>说&nbsp;&nbsp; 
                                          明：</font> 江湖等级高于<font color="#FFFF00"><b>200级</font>可以来这里得道飞升，经过得道飞升后你的等级将变为<font color="#FFFF00"><b>0级</font>，得道飞升后的体力，内力，武功上限都会有很大的提高，具体上限提升的计算为: 
                                          <font color="#FFFF00"><b>体力内力上限增加了你现在泡点总积分的十分之一（1/10），武功上限增加了你现在泡点总积分的十分之一（1/10）</font>！！<br>
                                          特别申明：攻击和防御只跟等级有关，所以得道飞升后你的攻击防御将变为1级时的数值（但是装备可以增加攻击）。
</font>
<br>
<br>
<br>
<div align="center"><a href="zsok.asp"><font color="#FFFF00" size=+2><b>我要脱胎换骨，得道飞升</font> 
                                          </div>
</td></form>
</tr>
</table>
<a href="#" onclick="javascript:window.close();"><img height=53 src="image/wonderofitallbutterflydark.jpg" width=60 border=0 alt="关闭窗口"></a><br>
<IMG height=25 src="image/wonderofitallbar.jpg" width=488 border=0> 
<br>
                            <font size="2" color="#FFFFFF">&copy; Copyright 2005 
                            www.7758530.com &#153; All Rights Reserved 快乐江湖 
                            ◆版权所有 禁止抄袭<br>
                            设计制作：回首当年　QQ:865240608</font> </TD>
                        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</TD></TR></TBODY></TABLE></CENTER></DIV></BODY></HTML>
