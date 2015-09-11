<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
search=mystr(server.HTMLEncode(Request("search")))
aqjh_name=session("aqjh_name")
%>
<html><title>我 的 江 湖 大 事 记</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<%
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if search<>"" then
rs.Open "select * from l where (b='"& aqjh_name &"' or c='"& aqjh_name &"') and d<>'管理记录' and (b like '%" & search & "%' or c like '%" & search & "%' or e like '%" & search & "%') Order by a DESC", conn, 1,1
else
rs.Open "Select * From l where (b='"& aqjh_name &"' or c='"& aqjh_name &"') and d<>'管理记录'  Order by a DESC", conn, 1,1
end if
if rs.eof and rs.bof then%>
<body>
<div align="center">对不起，暂时没有发现任何记录！！ <br>
[ <a href=# onclick=history.go(-1)>返回上一级</a> ]  </div>
<%else
RS.PageSize=12'设置每页记录数,可修改
Dim TotalPages
TotalPages = RS.PageCount
If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if
RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'设置最大记录数
Dim Totalcount
Totalcount =INT(RS.recordcount)
StartPageNum=1
do while StartPageNum+5<=CurPage
	StartPageNum=StartPageNum+5
Loop
EndPageNum=StartPageNum+4
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount%>
<table border="1" width="423" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a">
  <tr>
    <td colspan="5" align=center bgcolor="#d7ebff">我 的 江 湖 大 事 记</td>
</tr>
<tr>
<td  align=center >操作者</td>
<td align=center >被操作人</td>
<td  align=center >方式</td>
<td  align=center >操作数据</td>
</tr>
<%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (I<RS.PageSize)
p=p+1%>
<tr bgcolor="" onmouseover="this.bgColor='#85C2E0';" onmouseout="this.bgColor='';">
<td align=center><%=rs("b")%></td>
<td align=center><%=rs("c")%></td>
    <td align=center><%=rs("d")%></td>
<td align=left><a title="<%=rs("e")%>"><%if len(rs("e"))>30 then%><%=left(rs("e"),30)%>...<%else%><%=rs("e")%><%end if%></a></td>
</tr>
<%I=I+1
RS.MoveNext
Loop%>
<tr>
<td height="24" colspan=7 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr><form method="GET" action="jl.asp">
<td width="49%">关键字:<input name="fs" type="hidden" value="<%=fs%>" class="input1"> <input type="text" name="search" size=15 class="input1" onfocus=this.select() onmouseover=this.focus()><input name="B1" type="submit" value="搜索" class="input1"></td></form>
<td width="51%" align="right"> 有<font color=blue><%=Totalcount%></font>条 共<%=int(Totalcount/25)+1%>页 [
<% if StartPageNum>1 then %>
	<a href="jl.asp?CurPage=<%=StartPageNum-1%>&search=<%=search%>"><<</a>
<%end if%>
<%For I=StartPageNum to EndPageNum
if I<>CurPage then %>
	<a href="jl.asp?CurPage=<%=I%>&search=<%=search%>"><%=I%></a>
<%else %>
	<b><font color=#ff0000><%=I%></font></b>
<%end if %>
<%Next %>
<%if EndPageNum<RS.Pagecount then %>
	<a href="jl.asp?CurPage=<%=EndPageNum+1%>&search=<%=search%>">>></a>
<%end if%>
] </td>
</tr>
</table></td>
</tr>
</table>
</body></html>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
function mystr(restr)
restr=trim(restr)
restr=Replace(restr," ","")
restr=Replace(restr,"union","u_nion")
restr=Replace(restr,"select","s_elect")
restr=Replace(restr,"update","u_pdate")
restr=Replace(restr,"del","d_el")
restr=Replace(restr,"drop","d_rop")
restr=Replace(restr,"table","t_able")
restr=Replace(restr,";","；")
restr=Replace(restr,".","。")
restr=Replace(restr,"!","！")
restr=Replace(restr,"?","？")
restr=Replace(restr,"or","o_r")
restr=Replace(restr,"<","《")
restr=Replace(restr,">","》")
restr=Replace(restr,"\","、")
restr=Replace(restr,"/","、")
restr=Replace(restr,"'","")
restr=Replace(restr,chr(34),"")
restr=Replace(restr,"&","￡")
restr=Replace(restr,"#","＃")
restr=Replace(restr,"%","％")
restr=Replace(restr,chr(13) &chr(10),"<br>")
restr=Replace(restr,chr(13) ,"<br>")
restr=Replace(restr,chr(10) ,"")
mystr=restr
end function
%>