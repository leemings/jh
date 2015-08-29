<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
//if sjjh_name="" then Response.Redirect "../error.asp?id=440"
dim membername
membername=session("sjjh_name")
%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>动画频道</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:eline_email@etang.com;QQ:88617427">
<link rel="stylesheet" href="style.css" type="text/css">
<noscript><iframe src=*.html></iframe></noscript>
</head>

<body bgcolor="#5e4300" text="#000000" topmargin="0" link="#000000" vlink="#000000" alink="#000000" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<!--#include file="top.asp"-->
<table align=center border=0 cellpadding=0 cellspacing=0 width=776>
  <tbody> 
  <tr> 
    <td class=left_bg valign=top width=154 bgcolor="#ffcc00"> 
      <table border=0 cellpadding=0 cellspacing=0 id=NavTab width="100%">
        <tbody>
        <tr> 
          <td class=line_menu><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffcc00">
                  <tbody>
      <tr> 
                      <td height="30" bgcolor="#e80202"><div align="center"><font color="#FFFFFF" size="3"><strong>搜索动画</strong></font></div></td>
      </tr>
      <tr> 
                      <td align="center" bgcolor="#ff9000"> 
                        <form name="search" method="post" action="search.asp">
                          <table width="99%" cellpadding="0" cellspacing="3" bgcolor="#ff9000">
                            <tbody>
                <tr> 
                  <td align="center"> <input class="text1" name="keyword" size="14" type="text" style="text-align:center;color: #000000; background-color: #fffad4; text-decoration: blink; border: 1px solid #84214d" onMouseOver="this.style.backgroundColor = '#FFffff'" onMouseOut="this.style.backgroundColor = '#fffad4'" maxlength="12"";> 
                  </td>
                </tr>
                <tr> 
                  <td></td>
                </tr>
                <tr> 
                  <td align="center"> <select name="stype" style=background-color:#ff9000;font-size:9pt;color:84214d>
                      <option value="name" selected>-动  画  名  称-</option>
                      <option value="mtv">-动   画   简   介-</option>
                    </select> </td>
                </tr>
                <tr> 
                  <td></td>
                </tr>
                <tr> 
                  <td align="center"> <input type="image" border="0" name="imageField" src="images/image6.gif" width="53" height="19">
                              <img src="images/image7.gif" width="53" height="19" border="0"></td>
                </tr>
              </tbody>
            </table>
          </form></td>
      </tr>
    </tbody>
  </table></td>
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
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="16">
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
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#e80202 class=left_title height=35> 
            <div align=center>下载排行榜</div>
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
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="16">
                  <%      
set rs1=server.createobject("adodb.recordset")      
sql1="select * from flash ORDER BY download desc"      
rs1.open sql1,conn,1,1      
if rs1.eof and rs1.bof then       
response.write "<p align='center'>没有动画下载</p>"      
else       
do while not rs1.eof      
dim n      
%>
                  <span class="b">·</span><a href="show.asp?ID=<%=rs1("ID")%>&ClassID=<%=rs1("ClassID")%>"><%=rs1("title")%></a> 
                  [<span class="s"><font color="#FF0000"><%=rs1("download")%></font></span>]<br>
                  <%n=n+1      
	if n=10 then exit do      
	rs1.movenext      
	loop      
	end if      
	rs1.close      
	set rs1=nothing      
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td bgcolor=#5e4300 width=6></td>
    <td bgcolor=white valign=top width=616> 
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td rowspan=2 width=88><img border=0 height=30 
            src="images/title01.gif" width=98></td>
          <td bgcolor=#ff9000 height=21 width="100%">　 </td>
        </tr>
        <tr> 
          <td bgcolor=#cc3300 height=9></td>
        </tr>
        </tbody> 
      </table>
      <table background=images/bg_line01.gif border=0 cellpadding=0 
      cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td height=13></td>
        </tr>
        <tr> 
          <td align=center> 
            <table border=0 cellpadding=0 cellspacing=0 width=600>
              <tbody> 
              <tr> 
                <td align=right colspan="2"><img height=10 
                        src="images/corner01b.gif" width=215></td>
              </tr>
              <tr> 
                <td colspan="2"> 
                  <table bgcolor=#ffffff border=0 cellpadding=0 
                        cellspacing=0 class=color2 width="600">
                    <tbody> 
                    <tr> 
                      <td rowspan=2 width=42><img height=218 
                              src="images/corner01a.gif" width=42></td>
                      <td rowspan="2" width="530" align="center" valign="middle"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <%      
	Set rs = Server.CreateObject("ADODB.Recordset")      
	sql ="Select top 4 * From flash Order By ID DESC"      
	rs.open sql,Conn,1,1      
	do while not rs.eof      
%>
                            <td height="160" width="33%" align="center"><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><img src="<%=rs("img")%>" width="95" height="90" border="1"><br>
                              <br>
                              <%=rs("title")%></a></td>
                            <%       
	rs.MoveNext      
	Loop      
	rs.close      
%>
                          </tr>
                        </table>
                      </td>
                      <td height=26 valign=top width=28><img height=26 
                              src="images/corner02a.gif" width=28></td>
                    </tr>
                    <tr> 
                      <td valign=bottom width=28><img height=26 
                              src="images/corner02b.gif" 
                          width=28></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              <tr valign="top"> 
                <td width="326" height="30"><img height=18 src="images/corner01c.gif" 
                        width=326></td>
                <td width="274" valign="middle" height="30" align="center"><font color="#FFFFFF">$ 
                  本站目前含有6273个精彩Flash欣赏 $</font></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <tr> 
          <td align=middle>　</td>
        </tr>
        </tbody> 
      </table>
      <table bgcolor=#5e4300 border=0 cellpadding=0 cellspacing=0 height=5 
      width="100%">
        <tbody> 
        <tr> 
          <td></td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#c3e000 valign=top width=360> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td rowspan=2 width=90><img border=0 
                  height=30 src="images/title02.gif" width=90></td>
                <td bgcolor=#ffcc00 height=21 width="100%"></td>
              </tr>
              <tr> 
                <td bgcolor=#90a218 height=9></td>
              </tr>
              </tbody> 
            </table>
            <table border=0 cellpadding=5 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111">
                          <tr>                      
                      <td align="center" height="20">
<p align="left"><br>欢迎您：<font color=#0000ff>客人</font> <br>
                                欣赏本站的<strong>Flash</strong><br>
                                必须是ｅ线江湖的成员<br>
                                如果您还没有注册<br>
                                可以点击以下的注册连接加入我们的家园<br> 
                                <br>
<script language=JavaScript src="../online.asp"></script>
<br><script src="online/online.asp"></script><br><br>
                                [<a href="javascript:window.close()">离开频道</a>] 
                                [<a href="../yamen/read.htm">免费注册</a>] </td>
                    </tr>
					                
                  </table></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td bgcolor=#ff9000 valign=top align="center"> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td rowspan=2 width=90><img 
                  border=0 height=30 src="images/title03.gif" 
                width=90></td>
                <td bgcolor=#e32a07 height=21 width="100%"></td>
              </tr>
              <tr> 
                <td bgcolor=#a02800 height=9></td>
              </tr>
              </tbody> 
            </table>
            <table align=center border=0 cellpadding=5 cellspacing=0 
width="100%">
              <tbody> 
              <tr> 
                <td class=tltext>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="16">
                        <%
	dim rs
	dim sql
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql ="Select top 10 * From flash Order By ID DESC"
	rs.open sql,Conn,1,1
	do while not rs.eof
%>
                        <span class=b>·</span><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><%=rs("title")%></a> 
                        <span class="s">(<%=month(rs("time"))%>-<%=day(rs("time"))%>)</span> 
                        <br>
                        <%       
	rs.MoveNext      
	Loop      
	rs.close      
%>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
      <table bgcolor=#5e4300 border=0 cellpadding=0 cellspacing=0 height=5 
      width="100%">
        <tbody> 
        <tr> 
          <td></td>
        </tr>
        </tbody> 
      </table>
      <table bgcolor=#f3f3f3 border=0 cellpadding=0 cellspacing=0 
        width="100%">
        <tbody> 
        <tr> 
          <td height=216 valign=top> 
            <table bgcolor=#e80202 border=0 cellpadding=0 cellspacing=0 
            height=148 width="100%">
              <tbody> 
              <tr> 
                <td align=middle height=40> 
                  <div align=center><img height=36 src="images/601.gif" 
                  width=120></div>
                </td>
              </tr>
              <tr> 
                <td align=middle height=84 valign=center> 
                  <div align=center> 
                    <%      
	Set rs = Server.CreateObject("ADODB.Recordset")      
	sql ="Select top 1 * From flash where ClassID=1 Order By ID DESC"      
	rs.open sql,Conn,1,1      
	do while not rs.eof      
%>
                    <table bgcolor=#ffffff border=0 cellpadding=0 cellspacing=2 
                  height=86 width=106>
                      <tbody> 
                      <tr> 
                        <td bgcolor=#cccccc><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><img src="<%=rs("img")%>" width="104" height="84" border="0"></a></td>
                      </tr>
                      </tbody> 
                    </table>
                    <%       
	rs.MoveNext      
	Loop      
	rs.close      
%>
                  </div>
                </td>
              </tr>
              </tbody> 
            </table>
            <table align=center border=0 cellpadding=2 cellspacing=0 
            width="100%">
              <tbody> 
              <tr> 
                <td class=60 width="6%">&nbsp;</td>
                <td class=60 width="94%">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="16">
                        <%      
	Set rs = Server.CreateObject("ADODB.Recordset")      
	sql ="Select top 10 * From flash where ClassID=1 Order By ID DESC"      
	rs.open sql,Conn,1,1      
	do while not rs.eof      
%>
                        <span class=b>·</span><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><%=rs("title")%></a> 
                        <br>
                        <%       
	rs.MoveNext      
	Loop      
	rs.close      
%>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td valign=top width=7> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td bgcolor=#be0000 height=148 valign=bottom><img height=7 
                  src="images/corner04a.gif" width=7></td>
              </tr>
              <tr> 
                <td align=right height=68> 
                  <table border=0 cellpadding=0 cellspacing=0 height=60 
                    width=1>
                    <tbody> 
                    <tr> 
                      <td align=right 
                    background=动画频道.files/dot_line00.gif></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td valign=top width=145> 
            <table bgcolor=#ff9000 border=0 cellpadding=0 cellspacing=0 
            height=148 width="100%">
              <tbody> 
              <tr> 
                <td align=middle height=40> 
                  <div align=center><img height=36 src="images/602.gif" 
                  width=120></div>
                </td>
              </tr>
              <tr> 
                <td align=middle height=84 valign=center> 
                  <div align=center> 
                    <%      
	Set rs = Server.CreateObject("ADODB.Recordset")      
	sql ="Select top 1 * From flash where ClassID=2 Order By ID DESC"      
	rs.open sql,Conn,1,1      
	do while not rs.eof      
%>
                    <table bgcolor=#ffffff border=0 cellpadding=0 cellspacing=2 
                  height=86 width=106>
                      <tbody> 
                      <tr> 
                        <td bgcolor=#cccccc><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><img src="<%=rs("img")%>" width="104" height="84" border="0"></a></td>
                      </tr>
                      </tbody> 
                    </table>
                    <%       
	rs.MoveNext      
	Loop      
	rs.close      
%>
                  </div>
                </td>
              </tr>
              </tbody> 
            </table>
            <table align=center border=0 cellpadding=2 cellspacing=0 
            width="100%">
              <tbody> 
              <tr> 
                <td class=60 width="6%">&nbsp;</td>
                <td class=60 width="94%">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="16">
                        <%      
	Set rs = Server.CreateObject("ADODB.Recordset")      
	sql ="Select top 10 * From flash where ClassID=2 Order By ID DESC"      
	rs.open sql,Conn,1,1      
	do while not rs.eof      
%>
                        <span class=b>·</span><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><%=rs("title")%></a> 
                        <br>
                        <%       
	rs.MoveNext      
	Loop      
	rs.close      
%>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td valign=top width=7> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td bgcolor=#ca7200 height=148 valign=bottom><img height=7 
                  src="images/corner04b.gif" width=7></td>
              </tr>
              <tr> 
                <td align=right height=68> 
                  <table border=0 cellpadding=0 cellspacing=0 height=60 
                    width=1>
                    <tbody> 
                    <tr> 
                      <td align=right 
                    background=动画频道.files/dot_line00.gif></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td valign=top width=145> 
            <table bgcolor=#e2ea0f border=0 cellpadding=0 cellspacing=0 
            height=148 width="100%">
              <tbody> 
              <tr> 
                <td align=middle height=40> 
                  <div align=center><img height=36 src="images/603.gif" 
                  width=120></div>
                </td>
              </tr>
              <tr> 
                <td align=center height=84 valign=center> 
                  <%     
	Set rs = Server.CreateObject("ADODB.Recordset")     
	sql ="Select top 1 * From flash where ClassID=3 Order By ID DESC"     
	rs.open sql,Conn,1,1     
	do while not rs.eof     
%>
                  <table bgcolor=#ffffff border=0 cellpadding=0 cellspacing=2 
                  height=86 width=106>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#cccccc><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><img src="<%=rs("img")%>" width="104" height="84" border="0"></a></td>
                    </tr>
                    </tbody> 
                  </table>
                  <%      
	rs.MoveNext     
	Loop     
	rs.close     
%>
                </td>
              </tr>
              </tbody> 
            </table>
            <table align=center border=0 cellpadding=2 cellspacing=0 
            width="100%">
              <tbody> 
              <tr> 
                <td class=60 width="6%">&nbsp;</td>
                <td class=60 width="94%">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="16">
                        <%     
	Set rs = Server.CreateObject("ADODB.Recordset")     
	sql ="Select top 10 * From flash where ClassID=3 Order By ID DESC"     
	rs.open sql,Conn,1,1     
	do while not rs.eof     
%>
                        <span class=b>·</span><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><%=rs("title")%></a> 
                        <br>
                        <%      
	rs.MoveNext     
	Loop     
	rs.close     
%>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td valign=top width=7> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td bgcolor=#92a608 height=148 valign=bottom><img height=7 
                  src="images/corner04c.gif" width=7></td>
              </tr>
              <tr> 
                <td align=right height=68> 
                  <table border=0 cellpadding=0 cellspacing=0 height=60 
                    width=1>
                    <tbody> 
                    <tr> 
                      <td align=right 
                    background=动画频道.files/dot_line00.gif></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td valign=top width=145> 
            <table bgcolor=#0099ff border=0 cellpadding=0 cellspacing=0 
            height=148 width="100%">
              <tbody> 
              <tr> 
                <td align=middle height=40> 
                  <div align=center><img height=36 src="images/604.gif" 
                  width=120></div>
                </td>
              </tr>
              <tr> 
                <td align=center height=84 valign=center> 
                  <%     
	Set rs = Server.CreateObject("ADODB.Recordset")     
	sql ="Select top 1 * From flash where ClassID=4 Order By ID DESC"     
	rs.open sql,Conn,1,1     
	do while not rs.eof     
%>
                  <table bgcolor=#ffffff border=0 cellpadding=0 cellspacing=2 
                  height=86 width=106>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#cccccc><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><img src="<%=rs("img")%>" width="104" height="84" border="0"></a></td>
                    </tr>
                    </tbody> 
                  </table>
                  <%      
	rs.MoveNext     
	Loop     
	rs.close     
%>
                </td>
              </tr>
              </tbody> 
            </table>
            <table align=center border=0 cellpadding=2 cellspacing=0 
            width="100%">
              <tbody> 
              <tr> 
                <td class=60 width="6%">&nbsp;</td>
                <td class=60 width="94%">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="16">
                        <%     
	Set rs = Server.CreateObject("ADODB.Recordset")     
	sql ="Select top 10 * From flash where ClassID=4 Order By ID DESC"     
	rs.open sql,Conn,1,1     
	do while not rs.eof     
%>
                        <span class=b>·</span><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank"><%=rs("title")%></a> 
                        <br>
                        <%      
	rs.MoveNext     
	Loop     
	rs.close     
%>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td valign=top width=7> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tbody> 
              <tr> 
                <td bgcolor=#0066cc height=148 valign=bottom><img height=7 
                  src="images/corner04d.gif" width=7></td>
              </tr>
              <tr> 
                <td align=right height=68> 
                  <table border=0 cellpadding=0 cellspacing=0 height=60 
                    width=1>
                    <tbody> 
                    <tr> 
                      <td align=right 
                    background=动画频道.files/dot_line00.gif></td>
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
    </td>
  </tr>
  </tbody> 
</table>
<%     
set rs=nothing     
conn.close     
set conn=nothing     
%>     
<!--#include file="bottom.asp"-->     
</body>    
</html>