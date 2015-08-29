<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
stats="论坛搜索"
if boardmaster or master then
	guestlist=""
else
	guestlist=" lockboard<>1 and "
end if
call nav()
if BoardID="" or not isInteger(BoardID) then
	boardid=0
	call head_var(2,0,"","")
else
	boardid=cLng(Boardid)
	call head_var(1,BoardDepth,0,0)
end if
call main()
if founderr then call dvbbs_error()
call footer()

sub main()
	if Cint(GroupSetting(14))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛搜索的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
		founderr=true
		exit sub
	end if%>
<form action=queryresult.asp method=post>
<table cellpadding=5 cellspacing=1 align=center class=tableborder1>
	<tr>
		<th valign=middle colspan=2 >请输入要搜索的关键字</th>
	</tr>
	<tr>
		<td class=tablebody1 colspan=2 align=center valign=middle>(同时查询多个条件使用'<font  color=<%=Forum_body(8)%>>or</font>' 分隔关键字，查询同时满足某条件使用'<font  color=<%=Forum_body(8)%>>and</font>'分隔关键字)<br><br><input type=text size=40 name=keyword></td>
	</tr>
	<tr>
		<td class=tablebody2 valign=middle colspan=2 align=center><b>搜索选项</b></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>作者搜索</b>&nbsp;<input name=sType type=radio value=1></td>
		<td class=tablebody1 align=left valign=middle><select name=nSearch><option value=1>搜索主题作者<option value=2>搜索回复作者<option value=3>两者都搜索</select></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>关键字搜索</b>&nbsp;<input name=sType type=radio value=2 checked></td>
		<td class=tablebody1 align=left valign=middle><select name=pSearch><option value=1>在主题中搜索关键字<!--<option value=2>在帖子内容中搜索关键字<option value=3>两者都搜索--></select></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>日期范围</b></td>
		<td class=tablebody1 align=left valign=middle><select name=SearchDate class=smallsel><option value=ALL>所有日期<option value=1>昨天以来<option  selected value=5>5天以来<option value=10>10天以来<option value=30>30天以来</option></select><select name=Stable size=1><%
			For i=0 to ubound(AllPostTable)
				response.write "<option value="""&AllPostTable(i)&""""
				if trim(NowUseBBS)=trim(AllPostTable(i)) then response.write " selected "
				response.write ">"&AllPostTablename(i)&"</option>"
			next
		%></select></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>评分者搜索</b></td>
		<td class=tablebody1 align=left valign=middle><input type=text size=20 name=scoreuser></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>积分值搜索</b></td>
		<td class=tablebody1 align=left valign=middle><select name=SearchScore class=smallsel><option value=ALL selected>任意积分帖<option value=NONE>尚未评分帖<option value=A5>五分以上的加分帖<option value=A4>四分以上的加分帖<option value=A3>三分以上的加分帖<option value=A2>两分以上的加分帖<option value=A1>一分以上的加分帖<option value=Z0>零分帖<option value=B1>一分以上的减分帖<option value=B2>两分以上的减分帖<option value=B3>三分以上的减分帖<option value=B4>四分以上的减分帖<option value=B5>五分以上的减分帖</select></td>
	</tr>
	<tr>
		<td class=tablebody1 valign=middle colspan=2 align=center><b>请选择要搜索的论坛</b></td>
	</tr>
	<tr>
		<td class=tablebody1 colspan=2 valign=middle align=center><b>搜索论坛：</b>&nbsp;<select name=boardid size=1><option value=0>所有论坛</option><%
			set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
			do while not rs.EOF
				response.write "<option value="""&rs(0)&""" "
				if boardid=rs(0) then
					response.write "selected"
				end if
				response.write ">"
				select case rs(2)
				case 0
					response.write "╋"
				case 1
					response.write "&nbsp;&nbsp;├"
				end select
				if rs(2)>1 then
					for i=2 to rs(2)
						response.write "&nbsp;&nbsp;│"
					next
					response.write "&nbsp;&nbsp;├"
				end if
				response.write rs(1)&"</option>"
				rs.MoveNext
			loop
			set rs=nothing
		%></select></td>
	</tr>
	<tr>
		<td class=tablebody2 valign=middle colspan=2 align=center><input type=submit value=开始搜索></td>
	</tr>
</table>
</form>
<%call activeonline()
end sub%>