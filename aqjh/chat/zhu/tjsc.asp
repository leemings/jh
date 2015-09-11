<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.charset="gb2312"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
bkmn=Request.Form("money")
bkop=Request.Form("op")
id=Request.Form("money")
if id="" then
	Response.Write "<script language=JavaScript>{alert('提示：输入名字有错误！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if


if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：系统禁止字符，请确认你输入的名字是正确的！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 魅力,道德,转生 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：禁止捣乱');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if round(rs("道德"))<3000 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德不满3000');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("转生")<2 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：转生不到2次，没有这个权利');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
        rs.close

rs.open "select 组队,组名 FROM 用户 WHERE 姓名='"&bkmn&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&bkmn&"信息出错，请确认是否有此人');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("组队")=false  then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&bkmn&"没有打开组队状态！禁止加组！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

        rs.close

	set rs=nothing	
	conn.close
	set conn=nothing



select case bkop
case "踢出"
	Response.Redirect "jlzudui.asp?id=删除组员&to1="&bkmn&""
case "加组"
	Response.Redirect "jlzudui.asp?id=添加组员&to1="&bkmn&""
end select
%>
</body>
</html>