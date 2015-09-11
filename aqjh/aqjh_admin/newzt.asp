<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
mp=trim(request.form("mp"))
zm=trim(request.form("zm"))
xz=trim(request.form("xz"))
sm=trim(request.form("sm"))
xzsm=trim(request.form("xzsm"))
bds=trim(request.form("bds"))
if xz<>0 and xz<>1 then
	Response.Write "<script Language=Javascript>alert('提示：限制：0取消，1限制！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from p Where a='"&mp&"'",conn
if not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('提示：此门派已经存在！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "Select * from 用户 Where 姓名='"&zm&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('提示：未找到的掌门资料！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
conn.execute "update 用户 set 门派='" & mp & "', 身份='掌门',grade=5 where 姓名='"&zm &"'"
conn.execute "Insert Into p (a,b,c,d,e,f,g) Values('"&mp&"','"&zm&"',0,'"&sm&"','"&xzsm&"',"&xz&",'"&bds&"')"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','添加门派:"&request.form("subject")&"')"
rs.close
set rs=nothing
set rs=nothing
set bprs=nothing
Response.Write "<script Language=Javascript>alert('提示：门派添加完成！！');location.href = 'javascript:history.go(-1)';</script>"
%>

