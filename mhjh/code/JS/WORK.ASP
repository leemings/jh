<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<3 then 
s=3-sj 
Response.write "<script language='javascript'>alert ('我是剑侠之神，防止刷钱，所以无论你怎样刷\n都是无用的，请您等上["&s&"秒钟]再操作！');setTimeout('history.back();',1000);</script>"
Response.End 
end if 
Application.Lock 
Application("yx8_mhjh_dg")=now() 
Application.UnLock 
dj=session("yx8_mhjh_usergrade")
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from 打工 where id="&id,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=103"
ucondition=rst("条件")
uresult=rst("作用")
umoney=rst("银两")
rst.Close
rst.Open "select * from 用户 where 姓名='"&username&"' and 体力<"&dj*1000000&" and  "&ucondition,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=047"
rst.Close
set rst=nothing
conn.BeginTrans
conn.Execute "update 用户 set "&uresult&" where 姓名='"&username&"' and "&ucondition
conn.CommitTrans
conn.Close
set conn=nothing
Response.Write "<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel=stylesheet href='../style.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></head><body background='../image/bg.gif' oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>这里是"&umoney&"两银子，请收好了，同时您的体力增加了许多,体力超过上限(您的等级*1000000)时,您将无法继续补充!欢迎再来！<br>三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p></body>"
response.end
%>




