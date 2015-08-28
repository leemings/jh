<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
if instr(mg,"'") then Response.Redirect "../error.asp?id=024"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户  where 姓名='"&un&"' and 体力>100000 and 银两>100000",conn
if rst.EOF or rst.BOF then
   rst.Close
   conn.close
   Response.Write "<script language=JavaScript>{alert('您的银两连通关口都不够,"&mg&"拒绝你入境。带够10万银两和10万体力再来');}</script>"
   Response.Write "<script Language=Javascript>location.href = 'onlinelist.asp';</script>"
   Response.End
elseif rst("国家")=mg then
   rst.Close
   conn.close
   Response.Write "<script language=JavaScript>{alert('您已经是"&rst("国家")&"国的子民了，不用再加入');}</script>"
   Response.Write "<script Language=Javascript>location.href = 'onlinelist.asp';</script>"
   Response.End
elseif rst("积分")<4000 then
   rst.Close
   conn.close
   Response.Write "<script language=JavaScript>{alert('您的积分不够4000，不能加入');}</script>"
   Response.Write "<script Language=Javascript>location.href = 'onlinelist.asp';</script>"
   Response.End
else
conn.Execute "update 用户 set 国家='"&mg&"',职位='无',银两=银两-100000,体力=体力-100000,积分=积分-4000 where 姓名='"&un&"'"
joincountry="欢迎您加入"&mg
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=clng(Application("yx8_mhjh_talkpoint"))
dim newtalkarr(600)
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
newtalkarr(592)=2
newtalkarr(593)=0
newtalkarr(594)=un
newtalkarr(595)="大家"
newtalkarr(596)=""
newtalkarr(597)="#660099"
newtalkarr(598)="#660099"
newtalkarr(599)="<font color=FF0000>【加入】</font>##成功的加入"&mg&"国,积分减少4000,银两减少10万,体力减少10万<font class=timsty>（"&time()&"）<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
end if

rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title></title>
<link rel=stylesheet href='css3.css'>
<script language=javascript>
setTimeout("location.href='tb.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background="bg1.gif" >
<div align=center>
<hr noshade size="1" color=red>
三秒钟自动返回<br>
<input type=button value='返回' onclick="javascript:location.href='tb.asp'" id=button1 name=button1>
</div>
<p align="center"><%=joincountry%></p>
</body>
</html>
