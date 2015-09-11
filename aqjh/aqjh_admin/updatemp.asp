<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=trim(request.form("id"))
mp=trim(request.form("mp"))
zm=trim(request.form("zm"))
sm=trim(request.form("sm"))
xzsm=trim(request.form("xzsm"))
bds=trim(request.form("bds"))
jj=trim(request.form("jj"))
if mp="" or zm="" or sm="" or xzsm="" or bds="" or jj="" then
	Response.Write "<script Language=Javascript>alert('提示：不能有空的数据存在！！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from p Where id="&id,conn
yzm=rs("b")
rs.close
rs.open "Select * from p Where a='"&mp&"' and id<>"&id,conn
if not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('提示：此门派已经存在！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "Select * from 用户 Where grade<>10 and 姓名='"&zm&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('提示：未找到的掌门["&mp&"]资料！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
'取消原掌门
conn.execute "update 用户 set 门派='" & mp & "', 身份='弟子',grade=1 where 姓名='"& yzm &"'"
'设置新掌门
conn.execute "update 用户 set 门派='" & mp & "', 身份='掌门',grade=5 where 姓名='"&zm &"'"
'更新门派资料!
sqlstr="SELECT * FROM p where id="&id
rs.Open sqlstr,conn,1,2
rs("a")=mp
rs("b")=zm
rs("d")=sm
rs("e")=xzsm
rs("g")=bds
rs("h")=0
rs.Update
'conn.execute "update p set a='" & mp & "',b='" &zm& "',d='" &sm& "',e='" &xzsm& "',f=" &xz& ",g='"& bds &"',h="&jj&" where id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=701"
%>