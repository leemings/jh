

<%PageName="search"%>
<%on error resume next%>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<!--#include file="star.INC"-->
<!--#include file="top.asp"--> 
<%
keyword=trim(request("keyword")) 
keyword=replace(keyword,"'","''") 
stype=request("stype")
if keyword="" then
errmsg=errmsg+"�����ַ�����Ϊ�գ�����������ҵ���Ϣ<a href=""javascript:history.go(-1)"">�����ز�</a>"
call error()
Response.End 
end if
%><table cellpadding=0 cellspacing=0 border=0 width=100% align=center valign=middle> 
<tr>
<td align=center width="80%"> 
<table width="100%">
<tr>
      <td> <table width="766" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E7F3FF">
          <tr> 
            <td width="587" height="20">��<a href="index.asp"><font color="#000000">һ������ 
              &gt;</font></a> <font color="#000000">�����б�</font></td>
            <td width="179">��</td>
          </tr>
          <tr > 
            <td colspan="2" background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
          </tr>
        </table>
        <table width="766" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
          <tr>
            <td>
			
			
			
			
			
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
<td>
                    <div align="center">
                      <%
Set rs= Server.CreateObject("ADODB.Recordset")
if stype="Singer" then
sql="select * from Nclass where Nclass Like '%"& keyword &"%' order by Nclassid desc"
elseif stype="Music" then 
'sql="select * from MusicList where MusicName Like '%"& keyword &"%' order by id desc"
sql="select * from Musiclist where MusicName like '%"&keyword&"%'"
elseif stype="Special" then
sql="select * from Special where Name Like '%"& keyword &"%' order by Specialid desc"
else
end if
'response.write sql
'response.end
rs.open sql,conn,1,1

if not isempty(request.QueryString("page")) then
currentPage=cint(request.QueryString("page"))
else
currentPage=1
end if
if rs.eof and rs.bof then
if stype="Singer" then
response.write "<p align='center'><br><br>Sorry, δ �� �� �� �� Ҫ �� �� ��<br><br><a href=""javascript:history.go(-1)"">�� �� �� ��</a><br><br></p>"
elseif stype="Music" then 
response.write "<p align='center'><br><br>Sorry, δ �� �� �� �� Ҫ �� �� ��<br><br><a href=""javascript:history.go(-1)"">�� �� �� ��</a><br><br></p>"
elseif stype="Special" then
response.write "<p align='center'><br><br>Sorry, δ �� �� �� �� Ҫ �� ר ��<br><br><a href=""javascript:history.go(-1)"">�� �� �� ��</a><br><br></p>"
end if
else
totalPut=rs.recordcount
MaxPerPage=10
PageUrl="Search.asp"
if currentpage<1 then currentpage=1
if (currentpage-1)*MaxPerPage>totalput then
if (totalPut mod MaxPerPage)=0 then
currentpage= totalPut \ MaxPerPage
else
currentpage= totalPut \ MaxPerPage + 1
end if
end if
if currentPage=1 then
showpage totalput,MaxPerPage,PageUrl
showContent
showpage totalput,MaxPerPage,PageUrl
search
else
if (currentPage-1)*MaxPerPage<totalPut then
rs.move (currentPage-1)*MaxPerPage
dim bookmark
bookmark=rs.bookmark
showpage totalput,MaxPerPage,PageUrl
showContent
showpage totalput,MaxPerPage,PageUrl
search
else
currentPage=1
showpage totalput,MaxPerPage,PageUrl
showContent
showpage totalput,MaxPerPage,PageUrl
search
end if
end if
rs.close
end if

sub showContent 
i=0 
%>
                      <!-----------------------��������----------------------------->
                      <%if stype="Singer" then%>
                      <table border="1" width="88%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#56B0F4" bordercolordark="#FFFFFF">
<tr>
<td width="50%" height=22 align=center bgcolor="#B4DEF8">����</td>
<td width="25%" height=22 align=center bgcolor="#B4DEF8">��¼����</td>
<td width="25%" height=22 align=center bgcolor="#B4DEF8">��¼ר��</td>
</tr>
<%
Set Trs= Server.CreateObject("ADODB.Recordset")
do while not rs.eof
i=i+1
Tsql="SELECT * FROM MusicList where Nclassid="+cstr(rs("Nclassid"))
Trs.open Tsql,conn,1,1
TotalMNum=Trs.recordcount
Trs.close
Tsql="SELECT * FROM Special where Nclassid="+cstr(rs("Nclassid"))
Trs.open Tsql,conn,1,1
TotalSNum=Trs.recordcount
Trs.close
Tsql="SELECT * FROM Nclass where Nclassid="+cstr(rs("Nclassid"))
Trs.open Tsql,conn,1,1
Trs.close
%>
<tr>
<td width="50%">&nbsp;&nbsp;<%=i%>.&nbsp;<a href="Albumlist.asp?Name=һ������&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>"><%=rs("Nclass")%></a></td>
<td width="25%" align=center><%=TotalMNum%>��</td>
<td width="25%" align=center><%=TotalSNum%>��</td>
</tr>
<%
if i>=MaxPerPage then exit do
rs.movenext
loop
set Trs=nothing
%>
</table>

<!-----------------------��������----------------------------->
<%elseif stype="Music" then%>
<table border="1" width="90%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#56B0F4" bordercolordark="#FFFFFF">
<form name="form" onsubmit="javascript:return lbsong();" target="lbsong" action="Yxplaylist.asp"><tr>

<td width="7%" height=22 align=center bgcolor="#B4DEF8">ѡ��</td>
<td width="47%" height=22 align=center bgcolor="#B4DEF8">��������</td>
<td width="9%" height=22 align=center bgcolor="#B4DEF8">����</td>
<td width="9%" height=22 align=center bgcolor="#B4DEF8">Wma����</td>
<td width="9%" height=22 align=center bgcolor="#B4DEF8">���</td>
<td width="9%" height=22 align=center bgcolor="#B4DEF8">���ֺ�</td>
</tr>
<%do while not rs.eof%>
<tr>
<td align="center"><input type="checkbox" name="checked" value="<%=rs("id")%>"></td>
                            <td> <%=i%>.<a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','Ting88','scrollbars=no,resizable=no,width=518,height=355,menubar=no,top=168,left=168')" ><%=rs("MusicName")%></a></td>
<td align=center><%=rs("hits")%>��</td>
                            <td align=center>
                              <%if rs("Wma")<>"" then%>
                              <a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','Ting88','scrollbars=no,resizable=no,width=518,height=355,menubar=no,top=168,left=168')" >����</a> 
                              <%else%>
                              ���� 
                              <%end if%>
                            </td>
<td align=center><a href="#" onClick="window.open('Mailsong.asp?id=<%=rs("id")%>','','scrollbars=no,resizable=no,width=419,height=151,menubar=no,top=98,left=198')">����</a></td>
<td align=center><a href="Box.asp?action=add&id=<%=rs("id")%>" target="_blank" title="��Ա����ʹ�����ֺ�">����</a></td>
</tr>
<%
i=i+1
if i>=MaxPerPage then exit do
rs.movenext
loop
%>
<tr>
<td align="center" colspan="6" height="22">
<input type="button" name="chkall" value="ȫѡ" onclick="CheckAll(this.form)" title="ѡ����ʾ�����и���">&nbsp;
<input type="button" name="chkOthers" value="��ѡ" onclick="CheckOthers(this.form)" title="����ѡ�����">&nbsp;
<input type="submit" name="B1" value="����" title="����ѡ���������ĸ������ٵ������">
</td>
</tr>
</form>
</table>
<!-----------------------����ר��----------------------------->
<%elseif stype="Special" then%>
<table border="1" width="90%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#56B0F4" bordercolordark="#FFFFFF">
<tr>
<td width="30%" height=22 align=center bgcolor="#B4DEF8">ר��</td>
<td width="16%" height=22 align=center bgcolor="#B4DEF8">����</td>
<td width="17%" height=22 align=center bgcolor="#B4DEF8">��������</td>
<td width="9%" height=22 align=center bgcolor="#B4DEF8">����</td>
<td width="7%" height=22 align=center bgcolor="#B4DEF8">��¼����</td>
<td width="9%" height=22 align=center bgcolor="#B4DEF8">���</td>
</tr>
<%
Set Trs= Server.CreateObject("ADODB.Recordset")
do while not rs.eof
i=i+1
Tsql="SELECT * FROM MusicList where Specialid="+cstr(rs("Specialid"))
Trs.open Tsql,conn,1,1
TotalMNum=Trs.recordcount
Trs.close
%>
<tr>
<td width="30%">&nbsp;<%=i%>.<a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>"><%=rs("Name")%></a></td>
<td width="16%" align=center><a href="Albumlist.asp?Name=k666������&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>"><%=rs("Nclass")%></a>��</td>
<td width="17%" align=center><%=rs("Times")%>��</td>
<td width="9%" align=center><%=rs("Yuyan")%>��</td>
<td width="7%" align=center><%=TotalMNum%>��</td>
<td width="9%" align=center><%=rs("Hits")%>��</td>
</tr>
<%
if i>=MaxPerPage then exit do
rs.movenext
loop
Set Trs=nothing
%>
</table></div>
<%
else
end if
%>
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
dim n
if totalnumber mod maxperpage=0 then
n= totalnumber \ maxperpage
else
n= totalnumber \ maxperpage+1
end if
%>
<BR><div align="center">
<center>
<table width=95% border="1" cellpadding="2" cellspacing="0" bordercolor="#56B0F4" bordercolordark="#FFFFFF">
<tr><form method=Post action="<%=filename%>?keyword=<%=keyword%>&stype=<%=stype%>">
<td bgcolor="#B4DEF8" width="100%" align="center">���ҵ�<font color="<%=AlertFColor%>"><b><%=totalnumber%></b></font>���¼
&nbsp;��<strong><font color="<%=AlertFColor%>"><%=n%></font></strong>ҳ��ʾ
&nbsp;��ǰ��<strong><font color="<%=AlertFColor%>"><%=CurrentPage%></font></strong>ҳ 
<%if CurrentPage<2 then%>
&nbsp;��ҳ&nbsp;��һҳ&nbsp;
<%else%>
&nbsp<a href="<%=filename%>?page=1&keyword=<%=keyword%>&stype=<%=stype%>">��ҳ</a>&nbsp;
<a href="<%=filename%>?page=<%=CurrentPage-1%>&keyword=<%=keyword%>&stype=<%=stype%>"><font color="red"><b>��һҳ</b></font></a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
��һҳ&nbsp;ĩҳ
<%else%>
<a href="<%=filename%>?page=<%=CurrentPage+1%>&keyword=<%=keyword%>&stype=<%=stype%>"><font color="red"><b>��һҳ</b></font></a>&nbsp;
<a href="<%=filename%>?page=<%=n%>&keyword=<%=keyword%>&stype=<%=stype%>">ĩҳ</a>
<%end if%> 
</td></form> 

</tr>
</table>
</center>
</div><BR>
<% 
end function
function Search
%>
<%end function%>
</td> 
</table>
			
			
			
			
			
			
			
			
			</td>
          </tr>
        </table>
        </table>

</table>

<!--#include file="Bottom.asp"-->
<%
set rs=nothing
conn.close
set conn=nothing
%>