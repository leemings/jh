<%
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
ip=Request("ip")
if ip="" then Response.Redirect "../error.asp?id=060"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "SELECT * FROM 用户 where 最后登录ip='"&ip&"'",conn,1,3
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><p align=center><h4>按IP查询:<font color=FF0000>"&ip&"</font><h4></p><hr><table width=100% align=center border=3><tr bgcolor=FFFF00 align=center><td>姓名</td><td>密码</td><td>门派</td><td>等级</td><td>攻击</td><td>防御</td><td>体力</td><td>内力</td><td>银两</td><td>存款</td></tr>"
rst.PageSize=20
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
for j=1 to rst.Pagesize
	if rst.EOF or rst.BOF then
		exit for
	else
		Response.Write "<tr><td><a href='showuser.asp?search="&rst("姓名")&"'>"&rst("姓名")&"</a></td><td>"&rst("密码")&"</td><td>"&rst("门派")&"&nbsp;</td><td>"&rst("等级")&"</td><td>"&rst("攻击")&"</td><td>"&rst("防御")&"</td><td>"&rst("体力")&"</td><td>"&rst("内力")&"</td><td>"&rst("银两")&"</td><td>"&rst("存款")&"</td></tr>"
		rst.MoveNext
	end if
next	
Response.Write "</table><form action='showbyip.asp?ip="&ip&"' method=post id=form1 name=form1><table border=1 width=100% bgcolor=cccccc><tr align=center>"
if acpage>1 then
	response.write "<td><a href='showbyip.asp?ip="&ip&"&acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='showbyip.asp?ip="&ip&"&acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
	response.write "<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
	response.write "<td><a href='showbyip.asp?ip="&ip&"&acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='showbyip.asp?ip="&ip&"&acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
	response.write "<td>下一页</td><td>最后一页</td>"
end if
response.write "<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form></body>"
%>