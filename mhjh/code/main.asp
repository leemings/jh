<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("QUERY_STRING")
if instr(lastip,"or") then response.end
username=session("yx8_mhjh_username")
if username="" then response.redirect "../error.asp?id=016"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE><%=Application("yx8_mhjh_systemname")%></TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<STYLE type=text/css>BODY {
	
}
TD {
	COLOR: #e1e1e1; FONT-SIZE: 12px; LINE-HEIGHT: 18px
}
A {
	COLOR: #ffffff; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ffffff; TEXT-DECORATION: underline
}
</STYLE>
<BGSOUND loop=10 src="im/music3.mid">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</HEAD>
<BODY bgColor=#FFE4CA text=#99bea7 link=#0033CC leftMargin=10 topMargin=3 
MARGINWIDTH="0" MARGINHEIGHT="0" onunload="javascript:window.open('exit.asp','','height=1,width=1');" background="/images/bg1.gif">
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 height=480 width=751>
  <TBODY>
  <TR>
    <TD background="im/tj_07.jpg" vAlign=top 
width=763>
      <TABLE align=center border=0 cellPadding=8 cellSpacing=0 width=717 height="493">
          <TBODY>
        <TR>
              <TD width=766 height="493" vAlign=baseline> 
                <TABLE border=1 borderColor=#000000 cellPadding=5 cellSpacing=2 
            height=81 width=747>
                <TBODY>
                  <TR borderColor=#336666> 
                    <TD bgColor=#333333 height=75 width=174>��</TD>
                    <TD width=184 align="center" bgColor=#333333> <table width="75%"  border="0">
                        <tr> 
                          <td><a href="#" onclick="javascript:chatwin=window.open('chatroom/chatroom.asp','chat');" onmouseover="window.status='����������';return true;" onmouseout="window.status='';return true;" title="����������"><img src="IMG05.jpg" width="150" height="38" border="0"></a></td>
                        </tr>
                        <tr> 
                          <td><a href="EXIT.ASP" title="��ȷ��Ҫ��ʱ�뿪������" target="_parent" style="text-decoration: none"><font color="#FF0000" size="4" face="��������"> 
                            <script language=JScript src="onlinelist.asp"></script>
                            </font></a></td>
                        </tr>
                      </table></TD>
                    <TD width=36 height=75 align="center" bgColor=#333333><font color="#FFFF00">
                    <a target="_self" href="Exit.asp"><font color="#FFFF00">��              
                      ��</font><br>
                      <font color="#FF0000">����</font> <br>
                      <font color="#FFFF00">�� ��</font></a></font><a target="_self" href="Exit.asp">
                    </a> </TD>             
                    <TD bgColor=#333333 width=273>
                      <!--webbot bot="HTMLMarkup" startspan -->
                      <script language="JavaScript" src="news/news.asp"></script> 
                      <!--webbot bot="HTMLMarkup" endspan i-checksum="21367" -->
                    </TD>
                  </TR>
                </TBODY>
              </TABLE>
                <TABLE align=center border=0 cellPadding=0 cellSpacing=0 height=1 
            width="99%">
                  <TBODY>
                    <tr>
                      <TD width="47%" height="392" align=middle vAlign=top style="border-style: solid; border-width: 1"> 
                        <iframe name="I1" width="737" height="440" scrolling="no" border="0" frameborder="0" marginwidth="1" marginheight="1" src="11.htm">
                        �������֧��Ƕ��ʽ��ܣ�������Ϊ����ʾǶ��ʽ��ܡ�</iframe>
                        
                        </TD>
                    </tr>
                  </TBODY>
                </TABLE>
              </TD>
            </TR></TBODY></TABLE>
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=763>
        <TBODY>
        <TR>
              <TD height=70 align="left" valign="top" background="im/tj_07.jpg"> 
                <DIV align=center class=F16> <FONT color=#99bea7> Copyright &copy;            
                  2015 By <font color="#FF0000"><a href="http://yx8.net" target="_blank">���׵���</a></font><%if session("yx8_mhjh_username")=Application("yx8_mhjh_admin") then%>   
          <a target="_blank" href="gf/ll.asp">�ٸ�����</a>&nbsp;    
            <a href="chatroom/yskg.asp" target="_blank">������</a>
          <%end if%> <BR>          
                  ����ά�������׵��� QQ:865240608</font>           
                  <div id="Layer2" style="position: absolute; width: 200; height: 136; z-index: 2; left: -28; top: 82">
<!--#include file="cook.asp"-->
                  </div>
                </DIV></TD></TR></TBODY></TABLE></TR></TBODY></TABLE></HTML>