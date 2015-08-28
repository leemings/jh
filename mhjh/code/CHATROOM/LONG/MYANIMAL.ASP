<%@ LANGUAGE=VBScript%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>

<html>
<head>
<link rel=stylesheet href='../../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../bg1.gif' bgproperties="fixed" leftmargin="3" topmargin="0">
<div align="center"><br>
<font size="3" color="#FFFFFF"><br>
  </font><b>神    
  龙 窟</b><br>   
<!--#include file="data.asp"--><%  
sql= "SELECT * FROM myanimal where username='"&username&"' and rest=0"   
rs.open sql,connb,1,1   
if not(rs.EOF or rs.BOF) then    
%>   
  <table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">   
<%   
do while not (rs.eof or rs.bof)   
gong=rs("attack")   
outtime=rs("outtime")   
%>    
    <tr align="center">    
      <td colspan="2" height="42">龙名：<%=rs("animalname")%>-<%=rs("lei")%><br>   
      </td>   
    </tr>   
     
    <tr>    
      <td colspan="2">    
        <%if outtime=true then%>   
        <div align=center><a href=l.asp title="进入困龙窟">目前被因禁</a></div>   
        <%   
      else   
          %>   
        <div align="center"><a href=breed.asp?id=<%=rs("id")%> title="给你的神龙喂食！"><font color="#808000">养</font></a>    
          <font color="#FF0000"> </font>    
          <a href=modifymyanimalname.asp?id=<%=rs("id")%> title="给你的神龙取一个好听的名字"><font color="#808000">改</font></a>    
          <font color="#FF0000"> </font>    
          <a href=zengmyanimal.asp?id=<%=rs("id")%> title="将你的神龙送给你的朋友！"><font color="#808000">送</font></a>    
          <font color="#FF0000"> </font>    
          <a href="javascript:parent.talkfrm.settalk('//神灵','<%=rs("animalname")%>')" title=攻击：<%=gong%>><font color="#808000">攻</font></a><hr noshade size="1" color=darkred></div>   
      </td>   
    </tr>   
       
    <%   
    end if    
rs.movenext   
loop   
rs.Close   
set rs=nothing   
connb.close   
set connb=nothing   
%>    
  </table>   
<table width="100%" border="0" cellspacing="2" cellpadding="2">   
<tr>   
      <td>神龙可用来攻击敌人，当神龙状态不佳时，神龙将不听使唤</td>   
</tr>   
</table>   
<br>   
  <%   
else   
%>   
Sorry！您现在没有任何的神龙！    
  <%end if%>      
</div>      
</body>      
</html>
