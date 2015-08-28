<!--#include file="data2.asp"-->
<!--#INCLUDE file="check.asp"-->
<%
cname=check(request("nick"),"小羊名字") 
%>
<%
if session("yx8_mhjh_username")="" then
Response.Redirect "../../error.asp?id=016"
else

IF Request.ServerVariables("ReQuest_Method") = "POST" THEN
sheepname=request.form("nick")
if sheepname="" or len(sheepname)="" then
%>
<script language="Vbscript">
msgbox"小孩的名字填写不正确！",0,"提示"
history.back
</script>
<%
else
rs.open"select * from rules",conn,1,1
if rs.bof then
rs.close
conn.close
response.write"没有规则存在"
else
dellifeday=rs("dellifeday")
rs.close
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'",conn,1,1
if rs.bof then
rs.close
conn.close
%>
<script language="Vbscript">
msgbox"您不是这个小孩的主人！",0,"提示"
history.back
</script>
<%
else
logintoday=rs("logintoday")
feeddate=rs("feeddate")
life=rs("life")
hungry=rs("hungry")
clean=rs("clean")
sheephappy=rs("sheephappy")
sheephealth=rs("sheephealth")
rs.close
if datediff("d",logintoday,date)<>0 then
tempdatetoday=date
conn.execute"update sheep set logintoday='"&tempdatetoday&"' where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'"
temptime=datediff("d",feeddate,date)
templife=life-dellifeday*temptime
temphungry=hungry-dellifeday*temptime/2
if temphungry<0 then
temphungry=0
end if
tempclean=clean-dellifeday*temptime/2
if tempclean<0 then
tempclean=0
end if
tempsheephappy=sheephappy-dellifeday*temptime/2
if tempsheephappy<0 then
tempsheephappy=0
end if
tempsheephealth=sheephealth-dellifeday*temptime/2
if tempsheephealth<0 then
tempsheephealth=0
end if
'conn.execute"update sheep set clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"',hungry='"&temphungry&"' where sheepname='"&sheepname&"' and id="&tempid&" and user='"&session("yx8_mhjh_username")&"'"
conn.execute"update sheep set life='"&templife&"',hungry='"&temphungry&"',clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"' where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'"
end if
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'"
if rs("life")<=0 or rs("clean")<=0 or rs("sheephappy")<=0 or rs("sheephealth")<=0 or rs("hungry")<=0 then
rs.close
conn.execute"delete * from sheep where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'"
conn.close
%>
<script language="vbscript">
msgbox"由于您没有精心照顾小孩！您的小孩已经死了，重新再养一个吧。",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'",conn,1,1
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">


<HTML><HEAD><TITLE>江湖托儿所</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../style.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgcolor="000000" topMargin=0 oncontextmenu=self.event.returnValue=false>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=80%>
  <TBODY> 
  <TR>
    <TD height=158 vAlign=top width="100%"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=578>
        <TBODY> 
        <TR>
          <TD vAlign=top align=middle width=14 rowSpan=8><IMG      
            src="../../image/dragon.gif"><BR><IMG height=152      
            src="../../image/line.gif" width=1><BR><BR><IMG height=35      
            src="../../image/arrow.gif" width=10> </TD>     
          <TD vAlign=top width="517"><BR>
            <P class=p9><font color="#FF0000" size="2"><b>【高人提示】&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font>
                  </P>
                  <table border="0" width="103%" cellspacing="1" cellpadding="0"
                height="192">
                    <tr> 
                      <td class="p9" width="25%" height="18"><font color="#FFFF00">□-小孩生命值：</font></td>
                      <td class="p2" width="25%" height="18"> 
                        <font color="#FFFF00"> 
                        <%
if rs("life")>80 then
typelife="健康"
end if
if rs("life")<=80 and rs("life")>60 then
typelife="良好"
end if
if rs("life")<=60 and rs("life")>40 then
typelife="虚弱"
end if
if rs("life")<=40 and rs("life")>20 then
typelife="生病"
end if
if rs("life")<=20 and rs("life")>0 then
typelife="病危"
end if
if rs("life")<=0 then
typelife="死亡"
end if
response.write typelife
                      %>
                        </font>
                      </td>
                      <td class="p2" width="25%" height="18"><span class="p9"><font color="#FFFF00">□-小孩的生日</font></span><font color="#FFFF00">：</font></td>
                      <td class="p2" width="25%" height="18"><font color="#FFFF00"><%=rs("sheepdate")%></font></td>
                    </tr>
                    <tr> 
                      <td class="p3" width="25%" height="18"><span class="p9"><font color="#FFFF00">□-饥饿程度</font></span><font color="#FFFF00">：</font></td>
                      <td class="p3" width="25%" height="18"><font color="#FFFF00"><%=rs("hungry")%></font></td>
                      <td class="p3" width="25%" height="18"><span class="p9"><font color="#FFFF00">□-清洁程度</font></span><font color="#FFFF00">：</font></td>
                      <td class="p3" width="25%" height="18"><font color="#FFFF00"><%=rs("clean")%></font></td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="18"><span class="p9"><font color="#FFFF00">□-健康程度</font></span><font color="#FFFF00">：</font></td>
                      <td class="p2" width="25%" height="18"><font color="#FFFF00"><%=rs("sheephealth")%></font></td>
                      <td class="p9" width="25%" height="18"><font color="#FFFF00">□-快乐程度：</font></td>
                      <td class="p2" width="25%" height="18"><font color="#FFFF00"><%=rs("sheephappy")%></font></td>
                    </tr>
                    <tr> 
                      <td class="p9" width="100%" colspan="4" height="17"><font color="#FF0000">□-光荣妈妈敬告：</font></td>
                    </tr>
                    <tr> 
                      <td class="p9" width="100%" colspan="4" height="71"><font color="#FFFFFF">□-同的方法照顾你的小孩会得到不同的工作量，你只有长时间的积累经验，每次耐心细致的观察你的小孩每天的情况，才能养出打工的工作量很大的小孩来。你的小孩每天可管理三次(一周期),照顾小孩五个周期以上就可以到孤儿院打工了，现在每个工作单位的价格是1个金币。<br>
                        □-注：每次照料小孩花费50元，当小孩子的四项值中任何一项值为0时，小孩会死去。</font>
                        <p align="center">&nbsp;<a href="sellsheep.asp"><font color="#FFFF00">我已经没钱过日子了，唯有要卖掉儿子</font></a>
                        </p>
                      </td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepeat.asp">
                          <p align="left">
                            <input type="submit" value="喂奶" 
                        name="B1" style="font-family: 宋体; font-size: 12px">
                            <font                     
                        color="#E18C59">饥饿+10</font>      
                            <input type="hidden"                     
                        name="num" value="<%=tempid%>">       
                            <input type="hidden"                     
                        name="sheepname" value="<%=sheepname%>">       
                          </p>       
                        </form>       
                      </td>       
                      <td class="p2" width="25%" height="38">        
                        <form method="POST" action="sheepclean.asp">       
                          <p align="left">       
                            <input type="submit" value="洗澡"        
                        name="B12" style="font-family: 宋体; font-size: 12px">       
                            <font                     
                        color="#E18C59">清洁+10</font>      
                            <input type="hidden"                     
                        name="num" value="<%=tempid%>">       
                            <input type="hidden"                     
                        name="sheepname" value="<%=sheepname%>">       
                          </p>       
                        </form>       
                      </td>       
                      <td class="p2" width="25%" height="38">        
                        <form method="POST" action="sheepsun.asp">       
                          <p align="left">       
                            <input type="submit" value="阳光"        
                        name="B12" style="font-family: 宋体; font-size: 12px">       
                            <font                     
                        color="#E18C59">健康+10</font>      
                            <input type="hidden"                     
                        name="num" value="<%=tempid%>"> <input type="hidden"                    
                        name="sheepname" value="<%=sheepname%>">      
                          </p>      
                        </form>      
                      </td>      
                      <td class="p2" width="25%" height="38">       
                        <form method="POST" action="sheeppei.asp">      
                          <p align="left">      
                            <input type="submit" value="陪伴"       
                        name="B12" style="font-family: 宋体; font-size: 12px">      
                            <font                     
                        color="#E18C59">快乐+10</font>      
                            <input type="hidden"                     
                        name="num" value="<%=tempid%>">       
                            <input type="hidden"                     
                        name="sheepname" value="<%=sheepname%>">       
                          </p>       
                        </form>       
                      </td>       
                    </tr>       
                  </table>       
                  <P class=p9 align="center">        
                    <%if rs("feedsheepday")>=3 then%>       
                    <br>       
                  </P>       
            <div align="center">  
              <center>      
                  <table border="0" width="470" cellspacing="1" cellpadding="0"       
            height="60">      
                    <tr>       
                      <td class="p9" width="100%"><b><font color="#FF0000">恭喜！今天您的小孩在孤儿院打工了，打工时间为</font></b><%=rs("milk")%><b><font color="#FF0000">个工作单位</font></b><br>      
                        <font color="#000000"><b><font color="#FF0000">快到</font></b></font><a href="indexsheep.asp">孤儿院</a><b><font color="#FF0000">去工资吧！</font></b></td>      
                    </tr>      
                  </table>      
              </center>  
            </div>  
            <P align=center>&nbsp; </P>      
          </TD>          
          <TD vAlign=top align=middle width=37 rowSpan=8> </TD>       
          <TD vAlign=top align=middle width=17 rowSpan=8><IMG        
            src="../../image/dragon.gif"><BR><IMG height=152        
            src="../../image/line.gif" width=1><BR><BR><IMG height=35        
            src="../../image/arrow.gif" width=10> </TD>       
          <TD vAlign=top align=middle width=1 rowSpan=8> </TD>       
        </TR></TBODY></TABLE></TD></TR>      
  <TR>      
    <TD bgColor=#847939 height=1 width="100%"><IMG height=1 src="jh/page_point.gif"       
      width=1></TD>      
  </TR></TBODY></TABLE>      
<%end if                        
rs.close                        
conn.close                        
%> </BODY></HTML>      
<%                        
end if                        
end if                        
end if                        
end if                        
end if                        
end if                     
%>       
      
