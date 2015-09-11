<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
userid=Request.Form("id")
Response.Write "<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM 用户 where id="&userid
rs.Open sqlstr,conn,1,2
if Request.Form("submit")="新增" then rs.AddNew
if Request.Form("submit")="删除" then
sqlstr="delete * from 用户 where id="&userid
conn.Execute sqlstr
Response.Write "成功删除id="&userid&"的用户"
else
hy=0
dim nc(16)
nc(1)="w1"
nc(2)="w2"
nc(3)="w3"
nc(4)="w4"
nc(5)="w5"
nc(6)="w6"
nc(7)="w7"
nc(8)="w8"
nc(9)="w9"
nc(10)="w10"
nc(11)="z1"
nc(12)="z2"
nc(13)="z3"
nc(14)="z4"
nc(15)="z5"
nc(16)="z6"
for i=1 to rs.Fields.Count-25
myform=rs.Fields(i).Name
'检测
for j=1 to 16
 if myform=nc(j) then
	chkok="yes"
	exit for
 else
	chkok="no"
 end if
next
if Request.Form(i+1)="" and chkok="no" then
	Response.Write "<script Language=Javascript>alert('提示：["&rs.Fields(i).Name&"]的数据不能为空！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
'检测
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

if rs.Fields(i).Name="会员等级" then hy=clng(Request.Form(i+1))
if rs.Fields(i).Name="会员日期" and hy>0 then
	if DateDiff("d",date(),cdate(Request.Form(i+1)))<10 then
		Response.Write "<script Language=Javascript>alert('提示：["&rs.Fields(i).Name&"]会员日期不正确！当设\n置会员等级时，会员日期应大于10天！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
end if
if rs.Fields(i).Name="门派" then mp=cstr(Request.Form(i+1))
if rs.Fields(i).Name="身份" then sf=cstr(Request.Form(i+1))
if lcase(rs.Fields(i).Name)="grade" then dj=clng(Request.Form(i+1))
if dj>10 then
	Response.Write "<script Language=Javascript>alert('提示：管理员最高等级为10级！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
if dj<>0 then
	if mp="官府" and dj<6 then
		Response.Write "<script Language=Javascript>alert('提示：当门派为官府时，[管理级]应大于6！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if mp<>"官府" and dj>=6 then
		Response.Write "<script Language=Javascript>alert('提示：当门派不为官府时，[管理级]应小于6！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="掌门" and mp<>"官府" and dj<>5 then
		Response.Write "<script Language=Javascript>alert('提示：[身份]掌门的管理等级为5级！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="长老" and mp<>"官府" and dj<>4 then
		Response.Write "<script Language=Javascript>alert('提示：[身份]长老的管理等级为4级！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="护法" and mp<>"官府" and dj<>3 then
		Response.Write "<script Language=Javascript>alert('提示：[身份]护法的管理等级为3级！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="堂主" and mp<>"官府" and dj<>2 then
		Response.Write "<script Language=Javascript>alert('提示：[身份]堂主的管理等级为2级！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
end if
if rs.Fields(i).Name="论谈身份" and Request.Form(i+1)<>"/" and Request.Form(i+1)<>"版主" and Request.Form(i+1)<>"站长" then
		Response.Write "<script Language=Javascript>alert('提示：[论谈身份]只能为：/ 版主 站长，/表示由发贴数定！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
end if
if rs.Fields(i).Name="银两" then yin=clng(Request.Form(i+1))
if rs.Fields(i).Name="存款" then 
	yin=yin+clng(Request.Form(i+1))
	if yin>2000000000 then
		Response.Write "<script Language=Javascript>alert('提示：[银两、存款]总合不可以超过20亿！！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
end if
Response.Write rs.Fields(i).Name&"(<font color=blue>"&strtype&"</font>)："&rs.Fields(i).Value&"---->"&Request.Form(i+1)&"<font color=blue>(……资料正确!)</font></font><br><br>"
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
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','["&Request.Form("submit")&"]ID="&userid&"的记录！')"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<LINK href="css/css.css" type=text/css rel=stylesheet>
恭喜，用户资料修改完成！
<a href="fine.asp">返回</a>