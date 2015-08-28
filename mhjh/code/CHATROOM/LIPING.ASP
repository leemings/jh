<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
st=Request.QueryString("st")
if instr(mg,"'") or instr(st,"'") then Response.Redirect "../error.asp?id=024"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
if st="大家" then
usecurative="你不会是想对大家使用"&mg&"吧"
else
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 美貌,个性 from 物品 where 名称='"&mg&"' and 所有者='"&un&"' and 数量>0",conn
if rst.EOF or rst.BOF then
usecurative="你并没有"&mg&"可供使用"
else
adhp=rst("美貌")
admp=rst("个性")
if isnull(adespecial) then adespecial="无"
conn.execute "update 用户 set 美貌=美貌+"&adhp&",个性=个性+"&admp&" where 姓名='"&st&"'"
conn.execute "update 物品 set 数量=数量-1 where 名称='"&mg&"' and 所有者='"&un&"'"
msg="<font color=FF0000>【用药】</font>##对%%使用了"&mg&",使之美貌增加"&abs(adhp)&",个性增加"&admp
usecurative="你对"&st&"成功使用了"&mg
aberrantname=Application("yx8_mhjh_aberrantname")
usecurative="你对"&st&"成功使用了"&mg&aberranttxt
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
newtalkarr(593)=1
newtalkarr(594)=un
newtalkarr(595)=st
newtalkarr(596)=""
newtalkarr(597)="#660099"
newtalkarr(598)="#660099"
newtalkarr(599)="<font color=FF0000>【用药】<\/font>##对%%使用了"&mg&",使之美貌增加"&adhp&",个性增加"&admp&aberranttxt&"<font class=timsty>（"&time()&"）<\/font>"
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
end if
%>
<html>
<head>
<title>使用礼品</title>
<link rel=stylesheet href='css3.css'>
<script language=javascript>
setTimeout("location.href='seeliping.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background='bg1.gif'>
<div align=center>
<font color=0000ff size=3>使用礼品</font>
<hr>
三秒钟自动返回<br>
</div>

</body>
</html>
