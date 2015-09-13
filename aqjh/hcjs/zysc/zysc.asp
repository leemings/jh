<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML><HEAD><TITLE>自由市场</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../../css.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
<style type="text/css">
              <!--
              .glow{ filter: Glow(Color:#ff0000,Strength:1) }
              .glow1{ filter: Glow(Color:#ffff00,Strength:1) }
              -->
</style>
</HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=10 oncontextmenu=self.event.returnValue=false>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=596 height="376">
  <TBODY> 
  <TR>
    <TD vAlign=top width=596 height="25"> 
      <TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9 
      width="97%">
        <TBODY> 
        <TR>
          <TD height=25 width="31%">　</TD>
          <TD height=25 vAlign=top width="37%"> 
            <DIV align=center><font size="3" color="#FF0000">自由市场</font></DIV>
          </TD>
          <TD height=25 vAlign=top 
width="32%">　</TD>
        </TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD height=350 vAlign=top width="596"> 
      <TABLE align=center border=2 cellPadding=0 cellSpacing=0 class=mountain 
      width=596 height="347">
        <TBODY> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="345"><BR>
            <div align="center" style="width: 587; height: 328">
              <center>
            <TABLE border=0 
            cellPadding=0 cellSpacing=0 class=p9 width="550" style="border-collapse: collapse" bgcolor="#000000" bordercolor="#000000" height="1">
              <TBODY> 
              <TR >
                  <td  STYLE="background-color: transparent;color:#ffffff;border=0px none; " class="Glow" width="548" height="1">
                  <P class=p9  STYLE="background-color: transparent;color:#ffffff;border=0px none; " class="Glow" align="center">
                  <br>
                  这里是快乐江湖中最大的自由市场，在这里，你可以买到很多好东西，但没钱你也只能看看啊<br>
　</P>
                  </td STYLE="background-color: transparent;color:#ffffff;border=0px none; " class="Glow1">
                  <table width="548" border="0" cellspacing="0" cellpadding="0" align="center" height="224">
                    <tr> 
                     <td width="548" STYLE="background-color: #FFFFFF;color:#ffffff;border=0px none; " class="Glow" background="sc.jpg" height="224" align=center>
                       <b><a href="market.asp">
                       <font color="#FF0000">自由市场已经开门了，请进入吧．这里有你需要的东西！</font></a></b></td>
                    </tr>
                  </table>
                  <P align=center class="p9"><a href="javascript:window.close()">
                  <font color="#FFFF00">离开市场</font></a></P>
                  </TD></TR></TBODY></TABLE>
              </center>
          </div>
            </TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="596"></TD>
  </TR></TBODY></TABLE></BODY></HTML>