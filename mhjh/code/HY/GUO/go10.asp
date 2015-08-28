<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if Session("usepro")<>"5" then
Session("usepro")=""
response.redirect "index.asp"
end if
%>
<% 
if name="" then %>
<script language=vbscript>
MsgBox "对不起，你还没有登录或者超时"
location.href = "<a href="javascript:self.close()">"
</script>
<%end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+10000000,攻击=攻击+10000,防御=防御+500000,精力=精力+1000000 where 姓名='"&name&"'"
conn.Execute(sql)
Session("usepro")=""
conn.Close    
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=name 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【过关斩将】</font><font color=blue>"&name&"历经千辛万苦,终于闯过了5关,领悟到了一个江湖高手的真谛,并获得了巨额宝物!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<% 
if datediff("s",session("refresh"),now())<50000 then 
response.write "拒绝恶意刷新!小心删你帐号！" 
response.end 
else 
 session("refresh")=now()  

end if 
%>
<html>

<head>
<title>你终于取到了独孤求败的宝藏</title>
<LINK href="../../style.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background="bg.gif" leftmargin="0" oncontextmenu=self.event.returnValue=false >
<div align="center">
  <table border="0" width="600">
    <tr> 
      <td width="100%"> <p align="left" style="line-height: 200%; margin: 20"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <font color="#FFFF00">终于到达独孤求败的宝藏了，可是当你用内力震开宝藏的门是，却发现并没有什么天下无敌的武林秘籍，只是山洞的石壁上龙飞凤舞的写着这样一行字：</font></b></td>
    </tr>
    <tr> 
      <td width="100%"><img border="0" src="finish.gif"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="100%" valign="top"> 
        <p align="center" style="line-height: 200%; margin: 20"><b><font color="#000000">&nbsp;&nbsp;<font color="#FFFF00"> 
          原来独孤求败终生未求得一败的原因就是他的勇气和自信！原来我们的勇气和自信就是天下无敌的</font></font></b><font color="#FFFF00"><b>秘诀呀！（在山洞里找到</b></font><b><font color="#FFFFFF">1000000</font><font color="#FFFF00">两银子，受石壁上独孤求败真迹的启发，精力增加</font><font color="#FFFFFF">50000</font><font color="#FFFF00">点，攻击和防御各增加<font color="#000000">了</font></font><font color="#FFFFFF">10000</font><font color="#FFFF00">点！恭喜！！恭喜！！！）</font></b></p>
        </td>
    </tr>
  </table> 
</div> 
</html> 
