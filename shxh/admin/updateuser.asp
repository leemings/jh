<%
Response.Buffer =true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
userid=Request.Form("id")
if not isnumeric(userid) then Response.Redirect "../error.asp?id=024"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>用户管理</p><hr><table width=80% border=3 align=center><tr bgcolor=FFFF00><td width='35%'>属性</td><td>新值</td></tr>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM 用户 where id="&userid
rst.Open sqlstr,conn,1,2
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
if Request.Form("submit")=" 删 除 " then 
	mate=rst("配偶")
	username=rst("姓名")
	if mate<>"无" then conn.Execute "update 用户 set 配偶='无' where 姓名='"&mate&"'"
	conn.Execute "delete * from 信件 where 收信人='"&username&"'"
	conn.Execute "delete * from 物品 where 所有者='"&username&"'"
	conn.Execute "delete * from 媒婆 where 申请人='"&username&"'"
	conn.Execute "delete * from 攻击 where 姓名='"&username&"'"
	conn.Execute "delete * from bbs where 作者='"&username&"'"
	rst.Delete
	msg=msg&"<tr><td colspan=2>该记录已删除</td></tr>"
else
	for i=1 to rst.Fields.Count-1
		msg=msg&"<tr><td>"&rst.Fields(i).Name&"("&rst.Fields(i).Type&")</td><td>"&Request.Form(i+1)&"</td></tr>"
		select case rst.Fields(i).type
			case 202,130
				rst.Fields(i).Value=cstr(Request.Form(i+1))
			case 3
				rst.Fields(i).Value=clng(Request.Form(i+1))
			case 7 
				rst.Fields(i).Value=cdate(Request.Form(i+1))
			case 11
				rst.Fields(i).Value=cbool(Request.Form(i+1))	
			case else
				rst.Fields(i).Value=Request.Form(i+1)
		end select		
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
