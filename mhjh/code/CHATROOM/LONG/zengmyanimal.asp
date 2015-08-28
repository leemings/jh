<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
chatroomsn=Session("yx8_mhjh_userchatroomsn")
allonline=split(Application("yx8_mhjh_onlinename"&chatroomsn),";")
allonlineubd=ubound(allonline)
nickname=session("yx8_mhjh_username")
if nickname="" or session("yx8_mhjh_inchat")<>"in" then Response.Redirect "../../error.asp?id=016"
id=Request.QueryString ("id")
if id="" then Response.end
%>
<html>
<head>
<style></style>
<link rel=stylesheet href='../css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../bg1.gif' bgproperties="fixed" leftmargin="3">
<div align="center"><br>
 <span style='font-size:9pt'><b>转   
  送 神 龙</b></span><br> 
<!--#include file="data.asp"--><%  
rs.Open ("SELECT * FROM myanimal where username='"&nickname&"'and rest=0 and  id="&id),connb 
if not(rs.EOF or rs.BOF) then 
animalname=rs("animalname") 
%> 
<table cellpadding="3" cellspacing="1" border="0" align="center" width="100%"> 
<form method="post" action="zengmyanimalok.asp" id=form1 name=form1> 
<tr align="center">  
        <td colspan="2">  
          <input type="hidden" name="id" value="<%=id%>"> 
          <input type="text" readonly style="text-align:center;" name="animalname" size="10" maxlength="10" value="<%=animalname%>"> 
<br><br>  
选择好友：<br> 
<select name="towho" style="font-size:9pt"><%for i=1 to allonlineubd-1 
if allonline(i)<>Application("yx8_mhjh_admin") then 
Response.Write "<option value='"&allonline(i)&"'>"&allonline(i)&"</option>" 
end if 
next 
%></select> 
      <br> <br>
      <input type="submit" name="Submit" value="赠送"> 
</td> 
</tr> 
<% 
end if 
rs.Close 
set rs=nothing 
connb.close 
set connb=nothing 
%>  
    </form> 
  </table> 
   你可以将<%=animalname%>送给你的朋友,不过神兽因水土不服会降低一半的攻击力！ 
</div> 
<p align="center"><input type=button value='返回' onclick="javascript:location.href='myanimal.asp'"></p>
</body> 
</html> 
 
 
 
 

