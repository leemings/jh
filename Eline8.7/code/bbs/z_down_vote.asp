<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_down_conn.asp"-->
<!--#include file="inc/FORUM_CSS.asp"-->
<%dim CurPage,display,grade,content,downid
if request("id")="" then
	response.write "您没有选择相关软件，请返回"
	response.end
end if
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
	CurPage = 1
Else
	CurPage = CINT(Request.QueryString("CurPage"))
End If
display = CurPage
set rs=server.createobject("adodb.recordset")
if request("action")="save" then
	grade=request("grade")
	if request("content")<>"" then
		content=replace(replace(replace(request("content"),"'","‘"),"<","&lt;"),">","&gt;")
	end if
	downid=request("id")
	if session("truevote")<>downid then
		sql="select * from Dvote where (id is null)"
		rs.open sql,conndown,1,3
		rs.addnew
		if request("content")<>"" then
			rs("content")=content
		end if
		rs("grade")=grade
		rs("downid")=downid
		session("truevote")=downid
		rs.update
		rs.close
	else
		response.write "<script>alert('对不起！你已经进行了投票！');window.close();</script>"
	end if
end if
sql="select * from Dvote where downid="&request("id")&" order by id desc"
rs.open sql,conndown,1,1
dim v_num,pingrade
if rs.bof and rs.eof then
	V_num=0
	pingrade=0
else
	V_num=rs.recordcount
	pingrade=allgrade()/V_num
end if%>
<table border="0" width="100%" cellspacing="0" cellpadding="1">
	<tr>  
		<th align=center bgcolor=#CCCCFF height=20>为您喜爱的软件打分评论<br></th>
	</tr>
	<tr>  
		<td align=center>目前本软件共有<%=V_num%>人打分，<br>综合评分为<%=formatnumber(pingrade,1)%><br></td>
	</tr>
	<tr>
		<td align=center><form method=post action="?"><input type="hidden" name="action" value="save"><input name=grade type=radio value=1> 1 <input name=grade type=radio value=2> 2 <input checked name=grade type=radio value=3> 3 <input name=grade type=radio value=4> 4 <input name=grade type=radio value=5> 5 <input name=id type=hidden value='<%=Request("id")%>'><br>简短评论：<br><input name=content type=text size="30" class="smallinput" maxlength="100">&nbsp;<input type=submit value="提交" name="cmdok" class="buttonface"></form></td>
	</tr>
	<tr>  
		<th align=center height=20>已发表评论</th>
	</tr>
	<tr>  
		<td align=center><%
			if rs.eof and rs.bof then
				%>暂时没有文章<%
			else
				Dim Totalrec,n
				Totalrec = RS.RecordCount
				if Totalrec mod 10=0 then
					n= Totalrec \ 10
				else
					n= totalrec \ 10+1
				end if
				If CurPage>n Then CurPage=n
				if CurPage<1 then CurPage=1
				rs.PageSize = 10
				rs.AbsolutePage=curpage%>
				<table board=0 width=100%>
					<%i=0
					do while (Not RS.Eof) and (i<10)
						i=i+1%> 
						<tr>
							<td width=90%><font color="green">☉ </font><%=rs("content")%>&nbsp;&nbsp;[打分：<%=rs("grade")%>]</td>
						</tr>
						<%RS.MoveNext
					Loop%>
				</table> 
				<table width="100%" border="0" cellspacing="0" cellpadding="3" align="center">
					<tr>
						<td>页次： <font color="#CC0000"><%=CurPage%></font> / <%=n%></td>
						<td align="right">分页：<%call disppagenum(curpage,n,"?CurPage=","&id="&request("id"))%>&nbsp;&nbsp;</td>
					</tr>
				</table>
			<%end if
		%></td>
	</tr>
</table>
<%rs.close
set rs=nothing

function allgrade() 
	dim tmprs 
	tmprs=conndown.execute("Select sum(grade) As grades from Dvote where downid="&request("id")&"") 
	allgrade=tmprs("grades") 
	set tmprs=nothing 
	if isnull(allgrade) then allgrade=0 
end function%>