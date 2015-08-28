<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
if instr(mg,"'")<>0 then Response.Redirect "../error.asp?id=024"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 体力,内力,特效 from 物品 where 名称='"&mg&"' and 所有者='"&un&"' and 数量>0",conn
if rst.EOF or rst.BOF then
usecurative="你并没有"&mg&"可供使用"
else
adhp=0-rst("体力")
admp=rst("内力")
adespecial=mid(rst("特效"),2,2)
if isnull(adespecial) then adespecial="无"
conn.execute "update 用户 set 体力=体力-"&adhp&",内力=内力+"&admp&" where 姓名='"&un&"'"
conn.execute "update 物品 set 数量=数量-1 where 名称='"&mg&"' and 所有者='"&un&"'"
msg="<font color=FF0000>【用药】</font>##对%%使用了"&mg&",使之体力增加"&abs(adhp)&",内力增加"&admp
usecurative="你对"&st&"成功使用了"&mg
aberrantname=Application("yx8_mhjh_aberrantname")
if instr(aberrantname,";"&st&";")<>0 and adespecial<>"无" then
aberrantlist=Application("yx8_mhjh_aberrantlist")
aberrantlistubd=ubound(aberrantlist)
for i=1 to aberrantlistubd step 4
if aberrantlist(i)=st and aberrantlist(i+2)=adespecial then aberrantlist(i+3)=dateadd("s",adhp,aberrantlist(i+3))
next
Application.Lock
Application("yx8_mhjh_aberrantlist")=aberrantlist
Application.UnLock
erase aberrantlist
aberranttxt="，并使之"&adespecial&"程度减轻"&abs(adhp)
end if
usecurative="你成功使用了"&mg&aberranttxt

end if
rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>使用药品</title>
<link rel=stylesheet href='../style.css'>
<script language=javascript>
setTimeout("location.href='seeshiping.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<div align=center>
<b>
<font color="#000000" size="4" face="隶书">使用通知</font>
</b>
<hr>
三秒钟自动返回<br>
<input type=button value='返回' onclick="javascript:location.href='seeshiping.asp'">
</div>
<p align="center"><%=usecurative%></p>
      <SCRIPT language=JavaScript>                                                                                                                                                           
document.write("<img src='http://www.21ex.com/count/count.asp?Referer=&Width="+escape(screen.width)+"&Height="+escape(screen.height)+"' border=0 width=1 height=1>");                                                                               
document.write("</a>");                                                                               
</SCRIPT>    
</body>
</html>
