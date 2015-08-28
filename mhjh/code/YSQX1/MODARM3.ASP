<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
schid=Request.Form("schid")
if not isnumeric(schid) then Response.Redirect "../error.asp?id=024"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center>武器管理</p><hr><table width=80% border=3 align=center><tr bgcolor=FFFF00><td width='35%'>属性</td><td>新值</td></tr>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("yx8_mhjh_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM 商店 where id="&schid
rst.Open sqlstr,conn,1,2
if Request.Form("submit")="新增" then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
if Request.Form("submit")="删除" then
rst.Delete
msg=msg&"<tr><td colspan=2>该记录已删除</td></tr>"
else
for i=1 to rst.Fields.Count-1
msg=msg&"<tr><td>"&rst.Fields(i).Name&"("&rst.Fields(i).Type&")</td><td>"&Request.Form(i+1)&"</td></tr>"
if rst.Fields(i).type=202 then
rst.Fields(i).Value=cstr(Request.Form(i+1))
elseif rst.Fields(i).type=3 then
rst.Fields(i).Value=clng(Request.Form(i+1))
elseif rst.Fields(i).Type=7 then
rst.Fields(i).Value=cdate(Request.Form(i+1))
end if
next
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table><p align=center><input type='button' value='　确 定　' onclick='javascript:history.go(-2)'></p></body>"
Response.Write msg
%>
