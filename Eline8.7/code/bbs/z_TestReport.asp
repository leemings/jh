<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="Z_TestConn.ASP"-->
<%
'=========================================================
' Script Written by 绿水青山
'=========================================================
dim ID
stats="开心辞典 举报有问题的题目"

if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关题目。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的题目参数。"
else
	ID=request("id")
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请登录后进行相关操作。"
end if
if founderr then
	call nav()
	call head_var(0,0,"开心辞典","Z_test.asp")
	call test_err()
else
	call nav()
	call head_var(0,0,"开心辞典","Z_test.asp")
	if request("action")="send" then
		call send()
	else
		call main()
	end if
	call activeonline()
	if founderr then call test_err()
end if
call footer()
'=========================================
sub main()
	dim master_2
 	set rs=conn.execute("select adduser from [admin] order by id")
 	if not(rs.bof and rs.eof) then
		do while not rs.eof 
			master_2=""&master_2&"<option value="""&rs(0)&""">"&rs(0)&"</option>&nbsp;"
			rs.movenext
		loop	
	else
		master_2="无管理员"
	end if
	rs.close
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <form action="?action=send&id=<%=id%>" method=post>
    <tr>
    	<th valign=middle colspan=2 height=25>举报有问题的题目给管理员</th>
	</tr>
    <tr>
    	<td class=tablebody1 width="40%"><b>报告发送给哪个管理员：</b></td>
    	<td class=tablebody1 width="60%"><select name=master size=1><%=master_2%></select></td>
    </tr>
	<tr>
    	<td class=tablebody1 valign="top"><b>消息内容：</b><br>请写清楚题目的问题所在<br>如果您认为答案不正确，请给出正确的答案，并给出原因<br>如果可以的话请写出可以找到正确答案的书籍等，方便管理员查证<br>经查证如果答案确实有错的话，举报者将会得到一定的奖励<br>非必要情况下不要使用这项功能</td>
    	<td class=tablebody1><textarea name="content" cols="55" rows="6">管理员，您好，由于如下原因，我向你报告这个有问题的题目：</textarea></td>
    </tr>
	<tr>
    	<td colspan=2  class=tablebody2 align=center><input type=submit value="发 送" name="Submit"></td>
	</tr>
	</form>
</table>
<%
end sub

sub send()
	dim body
	dim writer
	dim incept
	dim topic
	body=checkStr(request("content"))
	writer=membername
	incept=checkStr(request("master"))
	sql="select title from [test] where ID="&ID
	set rs=aconn.execute(sql)
	if not(rs.bof and rs.eof) then
		topic="举报有问题的题目"
		body=body & chr(10) &"---------------------------------------"& chr(10) &"举报的题目如下："& replace(rs(0),"'","＇") & chr(10) &"编辑这个题目：[URL=Z_TestEdit.asp?action=edit&Tid="&ID&"]Z_TestEdit.asp?action=edit&Tid="&ID&"[/URL]"
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的题目不存在</li>"
		exit sub
	end if
	rs.close
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept&"','"&membername&"','"&topic&"','"&body&"',Now(),0,1)"
	conn.execute(sql)
	sucmsg="<li>报告发送成功！<br><li>经查证如果答案确实有错的话，您将会得到一定的奖励"
	
	call suc()
	set rs=nothing
end sub
%>