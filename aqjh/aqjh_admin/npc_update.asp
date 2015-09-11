<%@ LANGUAGE=VBScript codepage ="936" %>
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
userid=Request.Form("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM npc where id="&userid
rs.Open sqlstr,conn,1,2
if Request.Form("submit")="新增" then rs.AddNew
if Request.Form("submit")="删除" then
sqlstr="delete * from npc where id="&userid
conn.Execute sqlstr
Response.Write "成功删除id="&userid&"的Npc"
else
hy=0
for i=1 to rs.Fields.Count-1
if Request.Form(i+1)="" then
	Response.Write "<script Language=Javascript>alert('提示：["&rs.Fields(i).Name&"]的数据不能为空！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
if rs.Fields(i).type=11 and lcase(Request.Form(i+1))<>"true" and lcase(Request.Form(i+1))<>"false" then 
	Response.Write "<script Language=Javascript>alert('提示：["&rs.Fields(i).Name&"]为逻辑型如：True或False！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
if rs.Fields(i).type=3 and (not isnumeric(Request.Form(i+1))) then 
Response.Write "<script Language=Javascript>alert('提示：["&rs.Fields(i).Name&"]请使用数字输入！');location.href = 'javascript:history.go(-1)';</script>"
Response.End 
end if
if rs.Fields(i).type=135 and (not isdate(Request.Form(i+1))) then 
Response.Write "<script Language=Javascript>alert('提示：["&rs.Fields(i).Name&"]日期数据类型不正确！');location.href = 'javascript:history.go(-1)';</script>"
Response.End 
end if
if rs.fields(i).Type=202 then strtype="字符"
if rs.fields(i).Type=3 then strtype="数字"
if rs.fields(i).Type=135 then strtype="日期"
if rs.fields(i).Type=11 then strtype="逻辑"
Response.Write "<font size=-1>"&rs.Fields(i).Name&"(<font color=blue>"&strtype&"</font>)："&rs.Fields(i).Value&"---->"&Request.Form(i+1)&"<font color=blue>(……资料正确!)</font></font><br><br>"
if rs.Fields(i).type=202 or rs.Fields(i).type=203 then
rs.Fields(i).Value=cstr(Request.Form(i+1))
elseif rs.Fields(i).type=3 then
rs.Fields(i).Value=clng(Request.Form(i+1))
elseif rs.Fields(i).Type=7 then
rs.Fields(i).Value=cdate(Request.Form(i+1))
end if	
next
rs.Update
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
恭喜，npc资料修改完成！
<a href="npc_list.asp">返回</a>