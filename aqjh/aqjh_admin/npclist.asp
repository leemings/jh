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
<html><title>NPC�б�</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><LINK href="css/css.css" rel=stylesheet>
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
{if(confirm("�˲�����Ҫɾ����NPC�����ȷ��ɾ������Ŀ��"))self.location.href=url;}
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
	rs.Open "select * from npc where (n���� like '%" & search & "%'or n��ʽ like '%" & search & "%' or n������ like '%" & search & "%' or n��Ʒ like '%" & search & "%') Order by n����", conn, 1,1
else
	rs.Open "Select * From npc Order by n����", conn, 1,1
end if
if rs.eof and rs.bof then%>
<body topmargin="0">
<div align="center">�Բ�����ʱû�з����κμ�¼���� <br>[ <a href=# onclick=history.go(-1)>������һ��</a> ] 
  <%else
RS.PageSize=20'����ÿҳ��¼��,���޸�
Dim TotalPages
TotalPages = RS.PageCount
If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if
RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'��������¼��
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
    <td colspan="12" align=center bgcolor="#d7ebff"><%=Application("aqjh_chatroomname")%>NPC�б�<br>
      <span class="style1">��������ϵͳ��NPC���ô����ò�Ҫ�������,���ĸ���ֵȫ�Ǹ��ݾ�����������.</span><br>
    </td>
  </tr>
  <tr> 
    <td align=center >NPC����</td>
    <td align=center >�Ա�</td>
    <td align=center >NPC����</td>
    <td align=center >����</td>
    <td align=center >����</td>
    <td align=center >����</td>
    <td align=center >�书</td>
    <td align=center >����</td>
    <td align=center >����</td>
    <td align=center >����</td>
    <td align=center >������</td>
    <td align=center ><a href="addnpc.asp"><font color="#0000FF">�����NPC</font></a></td>
  </tr>
  <%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (I<RS.PageSize)
p=p+1%>
  <tr bgcolor="" onmouseover="this.bgColor='#85C2E0';" onmouseout="this.bgColor='';"> 
    <td align=center><%=rs("n����")%></td>
    <td align=center><%=rs("n�Ա�")%></td>
    <td align=right><%=rs("n����")%></td>
    <td align=right><%=rs("n����")%></td>
    <td align=right><%=rs("n����")%></td>
    <td align=right><%=rs("n����")%></td>
    <td align=right><%=rs("n�书")%></td>
    <td align=right><%=rs("n����")%></td>
    <td align=right><%=rs("n����")%></td>
    <td align=center><%=rs("n����")%></td>
    <td align=right><%=rs("n������")%></td>
    <td align=center><a href="modinpc.asp?id=<%=rs("id")%>">�޸�</a> | <a href=Javascript:del("jhgl.asp?act=delnpc&npcname=<%=rs("n����")%>");>ɾ��</a></td>
  </tr>
  <%I=I+1
RS.MoveNext
Loop%>
  <tr> 
    <td colspan=14 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff"> 
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr> 
          <form method="GET" action="npclist.asp">
            <td width="50%"> ���ؼ��֣� 
              <input type="text" name="search" size=15 class="input1" onfocus=this.select() onmouseover=this.focus()> 
              <input name="B1" type="submit" value="����" class="input1"> </td>
          </form>
          <td width="50%" align="right"> ��<font color=blue><%=Totalcount%></font>�� 
            ��<%=int(Totalcount/25)+1%>ҳ [ 
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
