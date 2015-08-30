<!--#include file="conn.asp"-->
<!--#include file="Star.INC"-->
<!--#include file="function.asp"-->
<%
Specialid=request.QueryString("AlbumID")
set rs=server.createobject("adodb.recordset")
sql="select * from Special where Specialid="&cstr(Specialid)
rs.open sql,conn,1,3
if rs.EOF then
	errmsg= "<li>服务器出错！请联系管理员</li>"
	call error()
	Response.End 
else
	rs("hits")=rs("hits")+1
	rs.update
	hits=rs("hits")
	Classid=rs("Classid")
	SClassid=rs("SClassid")
	NClassid=rs("NClassid")
	Nclass=rs("Nclass")
	name=rs("name")
	Gongsi=rs("Gongsi")
	Yuyan=rs("Yuyan")
	pic=rs("pic")
	intro=rs("intro")
%>
<head>
<title><%=rs("Nclass")%>=><%=rs("Name")%></title>
</head>
<!--#include file="top.asp"-->
<table width="765" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#F79618">
  <tr> 
    <td width="586" height="20" bgcolor="#66CCFF">　<a href="index.asp"><font color="#000000">一线视听 
      &gt;</font></a><font color="#000000"><a href="Artlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>"><%=rs("SClass")%></a> 
      &gt; <a href="Albumlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>"><%=rs("Nclass")%></a> 
      &gt; <%=rs("Name")%></font></td>
    <td width="179" bgcolor="#66CCFF">　</td>
  </tr>
  <tr > 
    <td colspan="2" background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
  </tr>
</table>
<div align="center"> 
  <table width="765" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
    <tr> 
      <td width="129" valign="top" bgcolor="#E7F3FF"> 
        <!--#include file="leftmenu.asp" -->
      </td>
      <td width="1" bgcolor="#FF0000" background="map/fix2.gif"></td>
      <td width="100%" valign="top"> 
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <TD><IMG height=37 src="images/down_t.gif" 
          width=579></TD>
          </TR>
        <TR>
          <TD>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                    <TD width="2%"><IMG height=69 
                  src="images/down_t2.gif" width=124></TD>
                    <TD align=middle width="95%"><a href="http://www.happyjh.com" target="_blank"></a> 
                    </TD>
                  </TR></TBODY></TABLE></TD>
          <tr> 
            <td width="100%" colspan="2"><img border="0" src="images/xiaod.gif" width="1" height="1"></td>
          </tr>
          <tr > 
            <td width="100%" colspan="2"background="map/fix.gif"> <p><img border="0" src="images/xiaod.gif" width="1" height="1"></p></td>
          </tr>
          <tr> 
            <td width="100%" height="0" colspan="2"></td>
          </tr>
          <tr> 
            <td width="100%" colspan="2"> <div align="center"> 
                <center>
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr> 
                      <td width="170%" height="8"></td>
                    </tr>
                    <tr> 
                      <td width="100%" height="12"><TABLE style="BORDER-COLLAPSE: collapse" borderColor=#111111 
            cellSpacing=0 cellPadding=4 width="100%" border=0>
                          <TBODY>
                            <TR> 
                              <TD width="79%"> <div align="center"><br>
                                  <TABLE width="97%" border=0 align=center cellPadding=0 cellSpacing=0 bordercolor="#FF6633">
                                    <TBODY>
                                      <TR> 
                                        <TD> <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                                            <TBODY>
                                              <TR> 
                                                <TD width="11%"><IMG height=36 
                  src="images/down_ciba.gif" width=188></TD>
                                                <TD width="79%" 
                background="images/down_ciba_bg.gif"> <DIV align=center></DIV></TD>
                                                <TD width="10%"><IMG height=36 
                  src="images/down_ciba_logo.gif" 
              width=51></TD>
                                              </TR>
                                            </TBODY>
                                          </TABLE></TD>
                                      </TR>
                                      <TR> 
                                        <TD 
          style="BORDER-RIGHT: #c0c0c0 1px solid; BORDER-LEFT: #c0c0c0 1px solid" 
          vAlign=top bgColor=#f6f6f6 height=83> <TABLE cellSpacing=5 cellPadding=0 width="100%" border=0>
                                            <TBODY>
                                              <TR> 
                                                <TD vAlign=top width="70%" height=94> 
                                                  <TABLE cellSpacing=3 cellPadding=0 width="100%" border=0>
                                                    <TBODY>
                                                      <TR bgColor=#ffffff> 
                                                        <TD width="17%"> <DIV align=center>歌手姓名</DIV></TD>
                                                        <TD width="83%"><b><%=rs("NClass")%></b></TD>
                                                      </TR>
                                                      <TR bgColor=#ffffff> 
                                                        <TD width="17%"> <DIV align=center>专集名称</DIV></TD>
                                                        <TD width="83%"><b><%=name%></b></TD>
                                                      </TR>
                                                      <TR bgColor=#ffffff> 
                                                        <TD width="17%"> <DIV align=center>发行日期</DIV></TD>
                                                        <TD width="83%"><b><%=rs("Times")%></b></TD>
                                                      </TR>
                                                      <TR bgColor=#ffffff> 
                                                        <TD width="17%"> <DIV align=center>唱片公司</DIV></TD>
                                                        <TD width="83%"><b><%=Gongsi%></b></TD>
                                                      </TR>
                                                      <TR bgColor=#ffffff> 
                                                        <TD width="17%" height=32> 
                                                          <DIV align=center>语言种类</DIV></TD>
                                                        <TD width="83%" height=32><b><%=Yuyan%></b></TD>
                                                      </TR>
                                                      <TR bgColor=#ffffff>
                                                        <TD height=32>专集简介</TD>
                                                        <TD height=32> 
                                                          <%if intro="" then%>
                                                          <%else%>
                                                          <%=htmlencode3(intro)%> 
                                                          <%end if%>
                                                        </TD>
                                                      </TR>
                                                    </TBODY>
                                                  </TABLE></TD>
                                                <TD vAlign=top width="30%" bgColor=#e5e5e5 height=94> 
                                                  <TABLE cellSpacing=6 cellPadding=0 width="100%" border=0>
                                                    <TBODY>
                                                      <TR> 
                                                        <TD bgColor=#cccccc><B>专辑图片：</B></TD>
                                                      </TR>
                                                      <TR> 
                                                        <TD height="175"><div align="center"><img src="<%=pic%>" alt="专辑图片" width="116" height="120"></div></TD>
                                                      </TR>
                                                    </TBODY>
                                                  </TABLE></TD>
                                              </TR>
                                            </TBODY>
                                          </TABLE></TD>
                                      </TR>
                                      <TR> 
                                      </TR>
                                    </TBODY>
                                  </TABLE>
                                </div>
                                <div align="center"></div>
                                <div align="center"></div>
                                <div align="center"></div>
                                <div align="center"></div>
                                </TD>
                            </TR>
                          </TBODY>
                        </TABLE></td>
                    </tr>
                  </table>
                </center>
              </div>
              <%
end if
rs.close
%>
              <div align="center"> 
                <center>
                  <table width="557" height="0%" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" bordercolor="#000000" style="border-collapse: collapse" border="1">
                    <%
sql="select id, MusicName,hits,Wma from MusicList where Specialid="&cstr(Specialid)
rs.open sql,conn,1,1
if rs.EOF then
response.write "<div align=center>未收录歌曲</div>"
else
%>
                    <form name="form" onsubmit="javascript:return lbsong();" target="lbsong" action="Yxplaylist.asp">
                      <tr bgcolor="#ffc184"> 
                        <td height=37% align=center width="20"> &nbsp;选</td>
                        <td height=37% align=center width="197">歌曲名字</td>
                        <td height=37% align=center width="45"> MP3下载</td>
                        <td height=37% align=center width="30">人气</td>
                        <td height=37% align=center width="40">试听</td>
                        <td height=37% align=center width="40"> 歌词</td>
                        <td height=37% align=center width="40">点歌</td>
                        <td height=37% align=center width="45">音乐盒</td>
                      </tr>
                      <%
i=0
do while not rs.eof
i=i+1
%>
                      <tr bgcolor="#F7FBFF"> 
                        <td height="37%" align="center" width="20"> <input type="checkbox" name="checked" value="<%=rs("id")%>"> 
                        </td>
                        <td bgcolor="#F7FBFF" width="197"> 　<%=i%>.<a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','tt90','scrollbars=no,resizable=no,width=300,height=300,menubar=no,top=168,left=168')" ><%=rs("MusicName")%></a></td>
                        <td bgcolor="#F7FBFF" width="45"> <p align="center"> <a target="_blank" href="http://mp3search.baidu.com/wstsearch?tn=baidump3&ct=134217728&rn=&word=<%=rs("MusicName")%>"> 
                            下载</a></td>
                        <td align=center width="30"><%=rs("hits")%>　</td>
                        <td align=center width="40"> 
                          <%if rs("Wma")<>"" then%>
                          <a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','tt90','scrollbars=no,resizable=no,width=300,height=300,menubar=no,top=168,left=168')" ><img border="0" src="images/real.gif"></a> 
                          <%else%>
                          暂无 
                          <%end if%>
                        </td>
                        <td align=center width="40"> <a href="http://mp3search.baidu.com/wstsearch?tn=baidump3lyric&ct=150994944&rn=10&word=<%=rs("MusicName")%>" target="_blank"> 
                          <img border="0" src="images/gezhi.gif" alt="查询歌词"></a></td>
                        <td align=center width="40"><a href="#" onClick="window.open('Mailsong.asp?id=<%=rs("id")%>','','scrollbars=no,resizable=no,width=419,height=151,menubar=no,top=98,left=198')">送友</a></td>
                        <td align=center width="45"><a href="Box.asp?action=add&id=<%=rs("id")%>" target="_blank" title="会员才能使用音乐盒">放入</a></td>
                      </tr>
                      <% 
rs.movenext 
loop 
%>
                      <tr bgcolor="#FFFFFF"> 
                        <td align="center" colspan="8" height="26%" width="549"> 
                          <input type="button" name="chkall" value="全 选" onClick="CheckAll(this.form)" title="选择显示的所有歌曲" style="font-size: 9pt; border-style: solid; border-width: 1"> 
                          &nbsp; <input type="button" name="chkOthers" value="反 选" onClick="CheckOthers(this.form)" title="反向选择歌曲" style="font-size: 9pt; border-style: solid; border-width: 1"> 
                          &nbsp; <input type="submit" name="B1" value="播 放" title="请先选择你想听的歌曲后再点击播放" style="color: #FFFFFF; font-size: 9pt; border-style: solid; border-width: 1; background-color: #FF0000"> 
                        </td>
                      </tr>
                    </form>
                    <%
end if
rs.close
%>
                  </table>
                </center>
              </div></td>
          </tr>
          <tr> 
            <td width="100%" height="10" colspan="2"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</div>
<div align="center"> </div>
<%
set rs=nothing
conn.close
set conn=nothing
%><!--#include file="Bottom.asp"-->