<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>վ������</title>
<script src=js/js.js></script>
</head>
<%
keyword=trim(request.form("keyword"))
keyword=replace(keyword,"'","''")
stype=request.form("stype")
if keyword="" then
	response.write "<script language=javascript>alert('���ҹؼ��ֲ���Ϊ�գ��뷵���������룡');history.back(1);</script>"
	response.end
end if
%>
<body>
<%
'---------------------------search----------------------
Set rs= Server.CreateObject("ADODB.Recordset")
if stype="musicname" then 
	sql="select * from MusicDJ where MusicName Like '%"& keyword &"%' order by id desc"
elseif stype="djuser" then
	sql="select * from MusicDJ where djuser Like '%"& keyword &"%' order by id desc"
elseif stype="user" then
	sql="select * from user where username Like '%"& keyword &"%' order by id desc"
else
response.write "<script language=javascript>alert('��ѡ��Ҫ���������ͣ�');history.back(1);</script>"
end if
rs.open sql,conn,1,1

if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
if rs.eof and rs.bof then
    if stype="musicname" then 
		response.write "<script language=javascript>alert('û���ҵ���������������');history.back(1);</script>"
	elseif stype="User" then
		response.write "<script language=javascript>alert('û���ҵ��������Ļ�Ա��');history.back(1);</script>"
		end if
else
	totalPut=rs.recordcount
	MaxPerPage=10
	PageUrl="search.asp"
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
			rs.move  (currentPage-1)*MaxPerPage
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
<%if stype="musicname" then%><div align="center">
  <center>
                  <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" style="border-collapse: collapse" bordercolor="#999999" height="28">
                    <tr>
                      <td align=center bgcolor="#FFCC66" >��������</td>
                      <td align=center bgcolor="#FFCC66" >������</td>
                      <td align=center bgcolor="#FFCC66" >����</td>
                      <td align=center bgcolor="#FFCC66" >�ղ�</td>
                      <td align=center bgcolor="#FFCC66" >��������</td>
                      <td align=center bgcolor="#FFCC66" >���</td>
                    </tr>
<%
i=0
do while not rs.eof
	i=i+1
%>
                    <tr>
                      <td bgcolor="#EC9F00"><%=rs("musicname")%>
<%if Session("IsAdmin")=true then%>
<br>
<a href="cnadmin/songedit.asp?id=<%=rs("id")%>"><font color="#FF0000">��</font></a><font color="#FF0000">
                </font>
<a href="cnadmin/songsave.asp?act=del&id=<%=rs("id")%>"><font color="#FF0000">ɾ</font></a><font color="#FF0000">
                </font>
<%end if%>

                      </td>
                      <td bgcolor="#EC9F00">
                      <p align="center"><%=rs("djuser")%></td>
                      <td align=center bgcolor="#EC9F00">
<img src=images/radio2.gif alt=���������� style="cursor:hand" onclick="window.open('<%if rs("musictype")="RealPlayer" then%>djplay.asp<%elseif rs("musicType")="MediaPlayer" then%>djplay_media.asp<%else%>djplay_swf.asp<%end if%>?songid=<%=rs("id")%>','HeiRui_Studio_Player','width=276,height=343');">

</td>
                      <td align=center bgcolor="#EC9F00">
                      <img src=images/coll.gif style=cursor:hand alt='�ղش�����' onclick=javascript:open_window('UserCollect.asp?action=add&id=<%=rs("id")%>','usercollect','width=500,height=300')></td>
                      <td align=center bgcolor="#EC9F00"><%=rs("dateandtime")%>��</td>
                      <td align=center bgcolor="#EC9F00"><%=rs("hits")%>��</td>
                    </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
                  </table>

  </center>
</div>
<!-----------------------������������----------------------------->
<%elseif stype="djuser" then%><div align="center">
  <center>
                  <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" style="border-collapse: collapse" bordercolor="#999999" height="28">
                    <tr>
                      <td align=center bgcolor="#FFCC66" >������</td>
                      <td align=center bgcolor="#FFCC66" >��������</td>
                      <td align=center bgcolor="#FFCC66" >����</td>
                      <td align=center bgcolor="#FFCC66" >�ղ�</td>
                      <td align=center bgcolor="#FFCC66" >��������</td>
                      <td align=center bgcolor="#FFCC66" >���</td>
                    </tr>
<%
i=0
do while not rs.eof
	i=i+1
%>
                    <tr>
                      <td bgcolor="#EC9F00"><%=rs("djuser")%>��</td>
                      <td bgcolor="#EC9F00">
                      <p align="left"><%=rs("musicname")%>
<%if Session("IsAdmin")=true then%>
<br>
<a href="cnadmin/songedit.asp?id=<%=rs("id")%>"><font color="#FF0000">��</font></a><font color="#FF0000">
                </font>
<a href="cnadmin/songsave.asp?act=del&id=<%=rs("id")%>"><font color="#FF0000">ɾ</font></a><font color="#FF0000">
                </font>
<%end if%></td>
                      <td align=center bgcolor="#EC9F00">
<img src=images/radio2.gif alt=���������� style=cursor:hand onclick=javascript:open_window('DJPlay_RM.asp?songid=<%=rs("id")%>','DJplay_song','scrollbars=no,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,copyhistory=no,top=10,left=10,width=369,height=310')>

</td>
                      <td align=center bgcolor="#EC9F00">
                      <img src=images/coll.gif style=cursor:hand alt='�ղش�����' onclick=javascript:open_window('UserCollect.asp?action=add&id=<%=rs("id")%>','usercollect','width=500,height=300')></td>
                      <td align=center bgcolor="#EC9F00"><%=rs("dateandtime")%>��</td>
                      <td align=center bgcolor="#EC9F00"><%=rs("hits")%>��</td>
                    </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
                  </table>

  </center>
</div>

<!-----------------------������Ա----------------------------->
<%elseif stype="user" then%>
<div align="center">
  <center>
                  <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" style="border-collapse: collapse" bordercolor="#999999">
                    <tr>
                      <td width="32" height=22 align=center bgcolor="#FFCC66" >ID</td>
                      <td width="214" height=22 align=center bgcolor="#FFCC66" >��Ա</td>
                      <td width="43" height=22 align=center bgcolor="#FFCC66" >�Ա�</td>
                      <td width="218" height=22 align=center bgcolor="#FFCC66" >����</td>
                      <td width="71" height=22 align=center bgcolor="#FFCC66" >��ҳ</td>
                      <td width="147" height=22 align=center bgcolor="#FFCC66" >OICQ</td>
                    </tr>
<%
i=0
do while not rs.eof
	i=i+1
%>
                    <tr>
                      <td align=center width="32" bgcolor="#EC9F00"><%'=rs("id")%><%=totalPut-(i+MaxPerPage*(currentpage-1)-1)%>��</td>
                      <td align=center width="214" bgcolor="#EC9F00"><a onclick=javascript:window.open('UserInfo.asp?id=<%=rs("id")%>','UserInfo','scrollbars=no,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,copyhistory=no,top=10,left=10,width=300,height=200')  style=cursor:hand title="��Ա��ϸ����"><%=rs("UserName")%></a><%if DateValue(rs("regDate"))=date() then%> <img src="images/new.gif"><%end if%>
<%if Session("IsAdmin")=true then%>
<br>
<a href="cndmin/UserModify.asp?id=<%=rs("id")%>"><font color="#FF0000">��</font></a><font color="#FF0000">
                </font>
<a href="cnadmin/UserDel.asp?id=<%=rs("id")%>"><font color="#FF0000">ɾ</font></a><font color="#FF0000">
                </font>
<%end if%>
</td>
                      <td align=center width="43" bgcolor="#EC9F00"><%if rs("sex")=true then%>��<%else%><font color="red">Ů</font><%end if%></td>
                      <td align=center width="218" bgcolor="#EC9F00"><%=rs("Email")%>��</td>
                      <td align=center width="71" bgcolor="#EC9F00">
                      <a target="_blank" href="<%=rs("homepage")%>">��˲鿴</a>��</td>
                      <td align=center width="147" bgcolor="#EC9F00"><%=rs("oicq")%>��</td>
                    </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
                  </table>
  </center>
</div>
<%end if%>
<%end sub%>
</body>
<%
function showpage(totalnumber,maxperpage,filename)
dim n
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<form method=Post action="<%=filename%>?keyword=<%=keyword%>&stype=<%=stype%>">
  <center>
  <p>���ҵ�<font color="<%=AlertFColor%>"><b><%=totalnumber%></b></font>���¼
<%if CurrentPage<2 then%>
  &nbsp;��ҳ ��һҳ&nbsp;
<%else%>
  &nbsp<a href="<%=filename%>?page=1&keyword=<%=keyword%>&stype=<%=stype%>">��ҳ</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&keyword=<%=keyword%>&stype=<%=stype%>">��һҳ</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
  ��һҳ ĩҳ
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&keyword=<%=keyword%>&stype=<%=stype%>">��һҳ</a>
  <a href="<%=filename%>?page=<%=n%>&keyword=<%=keyword%>&stype=<%=stype%>">ĩҳ</a>
<%end if%>
  &nbsp;ҳ��:<strong><font color="<%=AlertFColor%>"><%=CurrentPage%>/<%=n%></font></strong>ҳ
  ת��:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
<%next%>   
  </select> </p>
</form>        
<%        
end function

function Search
%>
<%end function%>

</html>