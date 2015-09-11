<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
cz=trim(request.querystring("cz"))
hysm=Trim(Request.Form("hysm"))
hydz=Trim(Request.Form("hydz"))
if cz="进入" and Instr(hysm,";")=0 then
	Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，格式如下：\n广告1;广告2;广告3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if cz="滚动" and Instr(hysm,";")=0 then
	Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，格式如下：\n滚动1;滚动2;滚动3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if cz="pk值日" and hysm="" then
	Response.Write "<script Language=Javascript>alert('提示：数据不能为空!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if cz="滚动" and Instr(hydz,";")=0 then
	Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，格式如下：\n问1|答案1;问2|答案2;问3|答案3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if cz="滚动" and Instr(hydz,"|")=0 then
	Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，格式如下：\n问1|答案1;问2|答案2;问3|答案3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM sm where a='"&cz&"'"
rs.Open sqlstr,conn,1,2
rs("c")=hysm
if cz="广告" then
	rs("b")=clng(hydz)
else
	rs("d")=hydz
end if
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：资料修改成功!');location.href = 'ggsm.asp';</script>"
%>
