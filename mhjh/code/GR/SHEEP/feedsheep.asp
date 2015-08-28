<%
Response.Expires=0
if Session("yx8_mhjh_username")="" then Response.Redirect "../error.asp?id=016"
%>
<!--#include file="data1.asp"-->
<HTML><HEAD><TITLE>托儿所</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK 
href="../../style.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=10 oncontextmenu=self.event.returnValue=false>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=80% height="352">
  <TBODY> 
  <TR>
    <TD height=401 vAlign=top width="100%"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=85% bgcolor="#000000">
        <TBODY> 
        <TR>
          <TD vAlign=top width="100%">
            <p align="center"><b><font face="宋体" size="4" color="#FFFF00">托儿所</font>
            </b>
            <P class=p9><font color="#FF0000"><b>【高人提示】</b></font><br>
            　<font color="#000000">　</font><font color="#FFFF00">这儿是喂养小孩的地方，你小孩就住在这儿，看着她无忧无虑的生活着，你觉得在江湖里不再孤独。不过你每天得来这照顾她哦，不然她会死掉的哦
                  </font>
                  </P>
            <div align="center">
              <center>
                  <table border="0" width="476" cellspacing="1" cellpadding="0" height="32">
                    <tr> 
                      <td width="472" height="30"> 
                        <table border="0" bordercolor="#ffffff" cellpadding="0"
                cellspacing="1" width="470" height="1">
                          <tr> 
                            <td height="1"> 
                              <form method="POST" action="checksheep.asp">
                                <table border="0" width="100%" cellspacing="1"
                        cellpadding="0" height="22">
                                  <tr> 
                                    <td class="p2" width="100%" height="24"> 
                                      <%
rs.open"select sheepname from sheep where user='"&session("yx8_mhjh_username")&"'",conn,1,1
if rs.bof then
rs.close
response.write "您还没有领养小孩呢！快领养一个吧"
else
                              %>
                                      <p align="center"><span class="p9"><font color="#FFFF00">你的孩子的名字叫</font></span><font color="#FFFF00"><%=rs("sheepname")%> 
                                      </font><font color="#FFCC00"> 
                                        <input   
                              type="hidden" name="nick" size="20"  
                              style="font-family: Tahoma; font-size: 12px"  
                              value="<%=rs("sheepname")%>">  
                                        </font>  
                                    </td>  
                                  </tr>  
                                  <tr>   
                                    <td class="p3" width="100%" height="1">   
                                      <p align="center">  
                                        <input type="submit"  
                              value="进入" name="B1"  
                              style="font-family: 宋体; font-size: 12px">  
                                    </td>  
                                  </tr>  
                                </table>  
                              </form>  
                              <%   
rs.close   
end if   
conn.close   
                      %>  
                            </td>  
                          </tr>  
                        </table>  
                      </td>  
                    </tr>  
                  </table>  
              </center>
            </div>
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%">
                  <p align="right">   <!--#include file="data1.asp"-->            
                          <p><font color="#FFCC00"><span class="p9">算工资了：</span>             
                            <%rs.open("select * from sheep where user='"&session("yx8_mhjh_username")&"'"),conn,1,1                             
if  rs.bof then%>              
                            <span class="p9">快快养个小孩为你赚钱啦！</span>                   
                            <%else                             
 if rs("feedsheepday")>=3then%>              
                            </font></p>              
                          <form method="POST" action="sellmilk.asp">              
                            <p><span class="p9">你孩子为你在孤儿院里打了<%=rs("milk")%>工作单位。你想向孤儿院收多少小时的工钱呢？               
                              <input                 
          type="text" name="milk" size="9"                 
          style="font-family: 宋体; font-size: 12px">              
                              <input type="submit"                 
          value="收工钱了" name="B2"                 
          style="font-family: 宋体; font-size: 12px">              
                              <input type="hidden"                 
          name="sheepname" value="<%=rs("sheepname")%>">              
                              <%else%>              
                              你的孩子还没找到工作呢</span>。</p>              
                          </form>              
                          <%end if%>              
                                        
                            <%end if%>              
                </td>
              </tr>
            </table>
 
          </TD></TR></TBODY></TABLE></TD></TR>   
  <TR>   
    <TD bgColor=#847939 height=1 width="100%"><IMG height=1 src="jh/page_point.gif"    
      width=1></TD>   
  </TR></TBODY></TABLE></BODY></HTML>   
