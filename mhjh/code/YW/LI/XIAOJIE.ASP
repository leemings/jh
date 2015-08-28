<%my=session("yx8_mhjh_username")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>魔幻江湖-女性画廊</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../style4.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR><style fprolloverstyle TYPE="text/css"><!--A:link{text-decoration:none}A:visited{text-decoration:none} -->A:hover {color: #ff00ff}</style></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="a15.gif" text="#C0C0C0" oncontextmenu=self.event.returnValue=false>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=629 align="center">
  <TBODY> 
  <TR>
    <TD background=garden.jpg width="88">-</TD>
    <TD align=middle vAlign=top width=661>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="601" height="124">
          <TBODY>
            <TR> 
              <TD width="321" height="28"> <font size="2"><font color="#ffffff" size="2"><font color="#696969"><font color="#373737">|</font><font color="#505050">|</font>|</font><font color="#828282">|</font><font color="#9b9b9b">|</font><font color="#b4b4b4">|</font><font color="#cdcdcd">|</font><font color="#e6e6e6">|</font></font></font><font color="#00FF00">|[</font><font color="#00FF00" face="Abadi MT Condensed Light">魔幻江湖</font><font color="#00FF00"><img height="11" src="a12.gif" width="15">女性画廊]</font><font size="2">|<font color="#e6e6e6">|</font><font color="#cdcdcd">|</font><font color="#b4b4b4">|</font><font color="#9b9b9b">|</font><font color="#828282">|</font><font color="#696969">|<font color="#505050">|</font><font color="#373737">|</font></font></font></TD>
              <TD width="82" height="28"> <p align="center"><a href="index.asp"><font color="#C0C0C0">[返回]</font></a></p></TD>
              <TD width="77" height="28"> <p align="center"><a href="javascript:window.close()"><font color="#C0C0C0">[关闭]</font></a></p></TD>
              <TD width="77" height="54" rowspan="3"> <img border="0" src="m004.gif"></TD>
            </TR>
            <TR> 
              <TD width="418" height="13" colspan="3"> <p align="right"><font color="#FF0000"><img border="0" src="tie.jpg">&nbsp;&nbsp;&nbsp;</font></TD>
            </TR>
            <TR> 
              <TD width="418" height="13" colspan="3"> <font color="#FF0000">*</font><font color="#C0C0C0">体力-100,道德-1000，银两减相应数目，本服务可以提升你的内力。</font></TD>
            </TR>
            <TR> 
              <TD class=bg1 colspan=4  
          style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" vAlign=top  
            width=557 height="52"> <table border=1 align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#C0C0C0">
                  <tr bgcolor=#74E76D bordercolor="#000000"> 
                    <td height="15" bgcolor="#000000" bordercolor="#696969" width="14%"> 
                      <div align="center"><font color="#FF0000">女伴</font></div></td>
                    <td height="15" bgcolor="#000000" bordercolor="#696969" width="13%"> 
                      <div align="center"><font color="#FF0000">美貌</font></div></td>
                    <td height="15" bgcolor="#000000" bordercolor="#696969"> <div align="center"><font color="#FF0000">服务的价格</font></div></td>
                    <td height="15"  bgcolor="#000000" bordercolor="#696969"> 
                      <div align="center"><font color="#FF0000">被服务后增加的内力</font></div></td>
                    <td height="15"  bgcolor="#000000" bordercolor="#696969"> 
                      <div align="center"><font color="#FF0000">选择进退</font></div></td>
                    <td height="15"  bgcolor="#000000" bordercolor="#696969"> 
                      <div align="center"><font color="#FF0000">赎身</font></div></td>
                  </tr>
                  <!--#include file="jiu.asp"-->
                  <% 
sql="SELECT * FROM 名妓" 
Set Rs=connt.Execute(sql) 
do while not rs.bof and not rs.eof 
%>
                  <tr bgcolor=#DEAD63> 
                    <td bgcolor="#000000" width="14%"> <center>
                        <font color="#FFFF00"><%=rs("姓名")%> </font></center>
                    <td bgcolor="#000000" width="13%"> <div align="center"><font color="#FFFF00"><%=rs("美貌度")%> </font></div>
                    <td bgcolor="#000000" width="17%"> <div align="center"><font color="#FFFF00"><%=rs("美貌度")*1.5%> </font></div></td>
                    <td bgcolor="#000000" width="17%"> <div align="center"><font color="#FFFF00"><%=rs("美貌度")/5%> </font></div></td>
                    <td bgcolor="#000000" width="25%"> <center>
                        <font color="#00FF00"><a href=girl.asp?id=<%=rs("id")%>><span class="calen-curr">和你的女伴进入房间</span></a> 
                        </font></center></td>
                    <td bgcolor="#000000" width="14%"> <center>
                        <font color="#00FF00"><a href=shu.asp?id=<%=rs("id")%>><span class="calen-curr">赎身</span></a> 
                        </font></center></td>
                  </tr>
                  <% 
rs.movenext 
loop 
rs.close
set rs=nothing 
connt.close
set connt=nothing 
%>
                </table></TD>
            </TR>
            <TR> 
              <TD colspan=4 width="587" height="18"> <p align="center"><font color="#FF0000" face="Abadi MT Condensed Light">游戏吧</font></p></TD>
            </TR>
          </TBODY>
        </TABLE>         
    </TD>        
    <TD background=garden.jpg width="13"><img border="0" src="liangyq.jpg"><img border="0" src="linxp.jpg"><img border="0" src="linxr.jpg"><img border="0" src="liury.jpg"><img border="0" src="liw.jpg"><img border="0" src="zhangsy.jpg"><img border="0" src="chenhl.jpg"></TD></TR>        
  </TBODY></TABLE>        
</BODY></HTML>

