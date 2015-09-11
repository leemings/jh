<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../../error.asp?id=439"
id=clng(Request.form("id"))
bt=trim(request.form("bt"))
fbsj=request.form("sj")
gg=trim(request.form("gg"))
ck=request.form("ck")
lb=request.form("lb")
if bt="" then
	response.write "<script language=JavaScript>{alert('错误：请填写公告标题！。');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
if lb="" then
	response.write "<script language=JavaScript>{alert('错误：请选择类别！。');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
lbdata=split(lb,"|")
lbz=clng(lbdata(0))
lbmc=lbdata(1)
if gg="" then gg="无内容"
gg=replace(gg,"<br>","")
gg=replace(gg,chr(13),"<br>")
if isnull(fbsj) or fbsj="" then	fbsj=now()
if isnull(ck) or ck="" then ck=0
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from 公告 where id="&id,conn,2,2
if rs.EOF or rs.BOF then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=JavaScript>{alert('错误：找不到你要修改的公告！。');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
rs.close
conn.execute "update 公告 set 标题='"&bt&"',内容='"&gg&"',时间='"&fbsj&"',查看次数="&ck&",所属类别="&lbz&",类别名称='"&lbmc&"' where id="&id
set rs=nothing
conn.close
set conn=nothing
response.write "<script language=JavaScript>{alert('公告修改成功！。');location.href = 'ggadmin.asp';}</script>"
response.end
%>