<%
Response.Expires=0
bgcolor=Application("yx8_mhjh_backgroundcolor")
bgimage=Application("yx8_mhjh_backgroundimage")
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
if not(usergrade=50 and usercorp="官府") then Response.Redirect "../error.asp?id=046"
schid=Request.Form("schid")
question=Request.Form("question")
answer=Request.Form("answer")
ti=Request.Form("ti")
money=Request.Form("money")
qiang=cbool(Request.Form("qiang"))
if not isnumeric(schid) then Response.Redirect "../error.asp?id=024"
'for each element in Request.Form
'	if Request.Form(element)="" then Response.Redirect "../error.asp?id=102"
'next	
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>答题</p><hr>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("yx8_mhjh_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM 答题 where id="&schid
rst.Open sqlstr,conn,1,2
if Request.Form("submit")="新增" then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
if Request.Form("submit")="删除" then 
	rst.Delete
	msg=msg&"<tr><td colspan=2>该记录已删除</td></tr>"
else
        rst("问题")=question
	rst("答案")=answer
	rst("提供者")=ti
	rst("奖金")=money
	rst("抢答")=qiang
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"<p align=center><input type='button' value='　确 定　' onclick='javascript:history.go(-2)'></p></body>"
Response.Write msg
%>
