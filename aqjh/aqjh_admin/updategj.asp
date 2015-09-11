<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=trim(request.form("id"))
gj=trim(request.form("guo"))
jz=trim(request.form("jz"))
sm=trim(request.form("sm"))
xzsm=trim(request.form("xzsm"))
bds=trim(request.form("bds"))
jj=trim(request.form("jj"))
if gj="" or jz="" or sm="" or xzsm="" or bds="" or jj="" then
	Response.Write "<script Language=Javascript>alert('提示：不能有空的数据存在！！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from guo Where id="&id,conn
yjz=rs("b")
rs.close
rs.open "Select * from guo Where a='"&gj&"' and id<>"&id,conn
if not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('提示：此国已经存在！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "Select * from 用户 Where grade<>10 and 姓名='"&jz&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('提示：未找到的君主["&gj&"]资料！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
'取消原君主
conn.execute "update 用户 set 国家='" & gj & "', 职位='百姓' where 姓名='"& yjz &"'"
'设置新君主
conn.execute "update 用户 set 国家='" & gj & "', 职位='君主' where 姓名='"&jz &"'"
'更新国家资料!
sqlstr="SELECT * FROM p where id="&id
rs.Open sqlstr,conn,1,2
rs("a")=gj
rs("b")=jz
rs("d")=sm
rs("e")=xzsm
rs("g")=bds
rs("h")=0
rs.Update
'conn.execute "update p set a='" & gj & "',b='" &jz& "',d='" &sm& "',e='" &xzsm& "',f=" &xz& ",g='"& bds &"',h="&jj&" where id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=701"
%>