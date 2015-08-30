<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then 
	Response.Write "<script Language=Javascript>top.location.href='http://www.happyjh.com';alert('提示：对不起，您还没有登陆江湖！');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：想欣赏Flash请先进入江湖聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 银两,银币 FROM 用户 WHERE 姓名='"& session("sjjh_name") &"'",conn,1,3
if rs("银两")<100000 or rs("银币")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "身上没100000银两或银币不足1个，不能欣赏Flash！"
	response.end
end if
rs("银两")=rs("银两")-100000
rs("银币")=rs("银币")-1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>★快乐江湖总站-Flash频道★→观看动画</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000" topmargin="0">
<!--#include file="top.asp"-->
<table align=center border=0 cellpadding=0 cellspacing=0 width=778>
  <tbody> 
  <tr> 
    <td class=left_bg valign=top width=154> 
      <table border=0 cellpadding=0 cellspacing=0 id=NavTab width="100%">
        <tbody> 
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td bgcolor=#e80202 height=38> 
            <div align=center><a href="Index.asp"><img 
            border=0 height=30 src="images/lm.gif" 
width=150></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal id=C10463000010 valign=center> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=1"><img border=0 
            height=23 src="images/l_mu1.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=3"><img border=0 
            height=23 src="images/l_mu2.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=3"><img border=0 
            height=23 src="images/l_mu3.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=4"><img border=0 
            height=23 src="images/l_mu4.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=9"><img border=0 
            height=23 src="images/l_mu5.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        <tr> 
          <td class=normal> 
            <div align=center><a 
            href="FlashClass.asp?ClassID=10"><img border=0 
            height=23 src="images/l_mu6.gif" width=135></a></div>
          </td>
        </tr>
        <tr> 
          <td class=line_menu></td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#e80202 class=left_title height=35> 
            <div align=center>观看排行榜</div>
          </td>
        </tr>
        <tr> 
          <td class=left_line height=5></td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=6 cellspacing=0 class=left_text 
        width="100%">
        <tbody> 
        <tr> 
          <td class=line_h22> 
            <%      
dim rs1,sql1      
set rs1=server.createobject("adodb.recordset")      
sql1="select * from flash ORDER BY hits desc"      
rs1.open sql1,conn,1,1      
if rs1.eof and rs1.bof then       
response.write "<p align='center'>没有热点动画</p>"      
else       
do while not rs1.eof      
dim k      
%>
            <span class="b">·</span><a href="show.asp?ID=<%=rs1("ID")%>&ClassID=<%=rs1("ClassID")%>"><%=rs1("title")%></a> 
            [<span class="s"><font color="#FF0000"><%=rs1("hits")%></font></span>]<br>
            <%k=k+1      
	if k=10 then exit do      
	rs1.movenext      
	loop      
	end if      
	rs1.close      
%>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td bgcolor=#5e4300 width=6></td>
    <td bgcolor=#ef4204 valign=top width=618> 
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td rowspan=2 width=88><img border=0 height=30 
            src="images/list.gif" width=98></td>
          <td bgcolor=#ff9000 height=21 width="100%">　 </td>
        </tr>
        <tr> 
          <td bgcolor=#cc3300 height=9></td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=8 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td align=middle bgcolor=#ef4204> 
            <% 
set rs=server.createobject("adodb.recordset") 
sql="select * from flash where id="&request("id") 
rs.open sql,conn,1,1 
url=rs("url") 
%>
            <table border=0 cellpadding=0 cellspacing=0 width=545 align="center">
              <tbody> 
              <tr> 
                <td align=middle colspan=2 height=21 bgcolor="#cc0000"> 
                  <table bgcolor=#ffffff border=0 cellpadding=3 cellspacing=1 
                  width="100%">
                    <tbody> 
                    <tr> 
                      <td bgcolor=#ff9900 height="20"> 
                        <div align=center><font color="#FFffff"><%=rs("title")%></font></div>
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan=2 height=350> <object codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 width=440 height=320 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000>
                    <param name=movie value="<%=url%>">
                    <param name=quality value=high>
                    <embed src="<%=url%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width=440 height=320>
                      <%=url%> 
                    </embed> 
                  </object></td>
              </tr>
              <tr> 
                <td align=middle height=21 width=120><font color="#ffff00">　动画名称：</font></td>
                <td height=21 width=474><font color="#ffff00"><%=rs("title")%></font></td>
              </tr>
              <tr> 
                <td align=middle height=21 width=120><font color="#ffff00">　上传时间：</font></td>
                <td height=21 width=474><font color=#ffff00><span class=time><%=rs("time")%></span> 
                  大小：<span class=time><%=rs("KB")%> K</span> 观看：<span class=time><%=rs("hits")%></span> 
                  次 下载：<span class=time><%=rs("download")%></span> 次 评价：<%=rs("commend")%></font></td>
              </tr>
              <tr> 
                <td align=middle height=21 width=120><font color="#ffff00">　动画简介：</font></td>
                <td height=21 width=474><font color=#ffff00> 
                  <% if len (rs("content"))>30 then %>
                  <%= left (rs("content"),28)%>... 
                  <% else %>
                  <%=rs("content")%> 
                  <% end if %>
                  </font></td>
              </tr>
              <tr> 
                <td align=middle height=21 width=120><font color="#ffff00">　作者信箱：</font></td>
                <td height=21 width=474> 
                  <% if rs("email")="" then%>
                  <font color="#ffff00">没有该动画作者的信箱</font> 
                  <% else %>
                  <a href="mailto:<%=rs("email")%>"><%=rs("email")%></a> 
                  <% end if %>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan=2 height=21><img align=absMiddle 
                  height=16 src="images/img_1.gif" width=16> <a 
                  href="<%=rs("url")%>"><font 
                  color=#ffff00><strong>全屏欣赏</strong></font></a>　　　　<img 
                  align=absMiddle height=16 src="images/download.gif" 
                  width=16> <a href="download.asp?ID=<%=rs("ID")%>" 
                  target=_blank><font 
                  color=#ffff00><strong>下载存盘</strong></font></a></td>
              </tr>
              </tbody> 
            </table>
            
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<% 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
%> 
<!--#include file="bottom.asp"--> 
</body> 
</html> 
