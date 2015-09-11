<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if session("aqjh_adminok")<>true then response.end
search=cstr(server.HTMLEncode(Request.QueryString("search")))
%>
<html><title>NPC列表</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><LINK href="css/css.css" rel=stylesheet>
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>
</head>
<script language=JavaScript>
function del(url)
{if(confirm("此操作将要删除此NPC人物，您确定删除此项目吗？"))self.location.href=url;}
</script>
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
	rs.Open "select * from npc where (n姓名 like '%" & search & "%'or n招式 like '%" & search & "%' or n出场词 like '%" & search & "%' or n物品 like '%" & search & "%') Order by n经验", conn, 1,1
else
	rs.Open "Select * From npc Order by n经验", conn, 1,1
end if
if rs.eof and rs.bof then%>
<body topmargin="0">
<div align="center">对不起，暂时没有发现任何记录！！ <br>[ <a href=# onclick=history.go(-1)>返回上一级</a> ] 
  <%else
RS.PageSize=20'设置每页记录数,可修改
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
do while StartPageNum+10<=CurPage
StartPageNum=StartPageNum+10
Loop
EndPageNum=StartPageNum+9
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount%>
  <br>
</div>
<table border="1" width="100%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
  <tr> 
    <td colspan="12" align=center bgcolor="#d7ebff"><%=Application("aqjh_chatroomname")%>NPC列表<br>
      <span class="style1">对于我们系统的NPC设置大家最好不要随意更改,他的各项值全是根据经验计算出来的.</span><br>
    </td>
  </tr>
  <tr> 
    <td align=center >NPC名字</td>
    <td align=center >性别</td>
    <td align=center >NPC经验</td>
    <td align=center >银两</td>
    <td align=center >体力</td>
    <td align=center >内力</td>
    <td align=center >武功</td>
    <td align=center >攻击</td>
    <td align=center >防御</td>
    <td align=center >敌人</td>
    <td align=center >死次数</td>
    <td align=center ><a href="addnpc.asp"><font color="#0000FF">添加新NPC</font></a></td>
  </tr>
  <%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (I<RS.PageSize)
p=p+1%>
  <tr bgcolor="" onmouseover="this.bgColor='#85C2E0';" onmouseout="this.bgColor='';"> 
    <td align=center><%=rs("n姓名")%></td>
    <td align=center><%=rs("n性别")%></td>
    <td align=right><%=rs("n经验")%></td>
    <td align=right><%=rs("n银两")%></td>
    <td align=right><%=rs("n体力")%></td>
    <td align=right><%=rs("n内力")%></td>
    <td align=right><%=rs("n武功")%></td>
    <td align=right><%=rs("n攻击")%></td>
    <td align=right><%=rs("n防御")%></td>
    <td align=center><%=rs("n敌人")%></td>
    <td align=right><%=rs("n死次数")%></td>
    <td align=center><a href="modinpc.asp?id=<%=rs("id")%>">修改</a> | <a href=Javascript:del("jhgl.asp?act=delnpc&npcname=<%=rs("n姓名")%>");>删除</a></td>
  </tr>
  <%I=I+1
RS.MoveNext
Loop%>
  <tr> 
    <td colspan=14 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff"> 
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr> 
          <form method="GET" action="npclist.asp">
            <td width="50%"> 　关键字： 
              <input type="text" name="search" size=15 class="input1" onfocus=this.select() onmouseover=this.focus()> 
              <input name="B1" type="submit" value="搜索" class="input1"> </td>
          </form>
          <td width="50%" align="right"> 有<font color=blue><%=Totalcount%></font>条 
            共<%=int(Totalcount/25)+1%>页 [ 
            <% if StartPageNum>1 then %>
            <a href="npclist.asp?CurPage=<%=StartPageNum-1%>"><<</a> 
            <%end if%>
            <% For I=StartPageNum to EndPageNum
if I<>CurPage then %>
            <a href="npclist.asp?CurPage=<%=I%>"><%=I%></a> 
            <% else %>
            <b><font color=#ff0000><%=I%></font></b> 
            <% end if %>
            <% Next %>
            <% if EndPageNum<RS.Pagecount then %>
            <a href="npclist.asp?CurPage=<%=EndPageNum+1%>">>></a> 
            <%end if%>
            ] </td>
        </tr>
      </table></td>
  </tr>
</table></body></html>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
