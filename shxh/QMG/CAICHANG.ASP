<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML><HEAD><TITLE>菜市场</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../wupin/jh.files/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0 background="../IMAGES/BG.GIF">
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=608>
  <TBODY> 
  <TR>
    <TD background=../wupin/jh.files/history_top_bg.gif vAlign=top width=608> 
      <TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9 
      width="100%">
        <TBODY>
        <TR>
          <TD height=25 width="31%">&nbsp;</TD>
          <TD height=25 vAlign=top width="33%">
            <DIV align=center></DIV>
            <DIV align=center>　</DIV>
            <DIV align=center>　</DIV>
            <DIV align=center>　</DIV>
            <DIV align=center><b><font face="楷体_GB2312" color="#FF0000" size="5">江湖菜市场</font></b></DIV>
          </TD>
          <TD height=25 vAlign=top 
width="36%">&nbsp;</TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD background=../wupin/jh.files/history_table_bg.gif height=158 vAlign=top width="608"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=560>
        <TBODY> 
        <TR>
          <TD vAlign=top><BR>
            <TABLE align=center background=../wupin/jh.files/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=table1 width="100%">
              <TBODY> 
              <TR>
                <TD align=middle vAlign=center> 
                  <P class=p9 align="left"><font color="#FF0000"><b>【师傅提示】</b></font><br>
                  <font color="#000000">　　这儿是江湖中最大的菜市场，你可以在这买到做面用的主料，辅料和汤料，买了原料就块去江湖巧面馆做一份美味的巧面哦。</font><br>
                  </P>
                  <table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
                    <tr bgcolor="#FFFFFF"> 
                      <td height="16" colspan="5"> 
                        <p align=center class="p9"><font color="#FF3333">菜市场报价表</font></p>
                    <tr bgcolor=#74E76D bordercolor="#000000"> 
                      <td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="12%"> 
                        <div align="center"><font size="2" color="#FF0000">菜名</font></div>
                      </td>
                      <td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="18%"> 
                        <div align="center"><font size="2" color="#FF0000">类型</font></div>
                      </td>
                      <td height="15" width="16%" bgcolor="#66FF66" bordercolor="#FF6600"> 
                        <div align="center"><font size="2" color="#FF0000">增加美味度</font></div>
                      </td>
                      <td height="15" width="22%" bgcolor="#66FF66" bordercolor="#FF6600"> 
                        <div align="center"><font size="2" color="#FF0000">售价</font></div>
                      </td>
                      <td height="15" width="32%" bgcolor="#66FF66" bordercolor="#FF6600"> 
                        <div align="center"><font size="2" color="#FF0000">购买</font></div>
                      </td>
                    </tr>
                    <!--#include file="dadata.asp"--> <%
sql="SELECT * FROM 菜场 "
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%> 
                    <tr bgcolor=#DEAD63> 
                      <td bgcolor="#000000"> 
                        <center>
                          <font size="2"><%=rs("物品名")%><span class="p9"><font color="#6666FF"></font></span> 
                          </font> 
                        </center>
                        <font size="2"></font> 
                      <td bgcolor="#000000"> 
                        <div align="center"><font size="2"><%=rs("类型")%></font> 
                        </div>
                      
                      <td bgcolor="#000000"> 
                        <div align="center"><font size="2"><%=rs("美味度")%> </font></div>
                      </td>
                      <td bgcolor="#000000"> 
                        <div align="center"><font size="2"><%=rs("银两")%> </font></div>
                      </td>
                      <td bgcolor="#000000"> 
                        <center>
                          <font size="2"><a href=caichang1.asp?id=<%=rs("id")%>><span class="calen-curr"><font color="#FFFFFF">购买</font></span></a></font> 
                        </center>
                      </td>
                    </tr>
                    <%
rs.movenext
loop


set rs=nothing 

connt.close
%> 
                  </table>
                  <P align=center class="p9"><b><font color="#74E76D">【</font><a href="qmg.htm"><font color="#74E76D">返回面馆</font></a><font color="#74E76D">】</font></b></P>
                  <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE>
            <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="608"><IMG height=1 src="../wupin/jh.files/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>
