<%PageName="SongList"%>
<!--#include file="conn.asp"-->
<!--#include file="Star.INC"-->
<!--#include file="Top.asp"-->
<%
if request("page")<>"" then
	currentPage=cint(request("page"))
else
	currentPage=1
end if

if request("Name")<>"" then
	classid=request("Name")
else
	classid=""
end if
if request("ID")<>"" then
	Sclassid=request("ID")
else
	Sclassid=""
end if
if request("ArtID")<>"" then
	Nclassid=request("ArtID")
else
	Nclassid=""
end if
if request("AlbumID")<>"" then
	Specialid=request("AlbumID")
else
	Specialid=""
end if
%>
<%
set rs=server.createobject("adodb.recordset")
if classid<>"" then
	if Specialid<>"" then
		sql="select * from MusicList where Specialid="+cstr(Specialid)+" and Nclassid="+cstr(Nclassid)+" and sclassid="+cstr(sclassid)+" and classid="+cstr(classid)+"  order by id desc" 
	elseif Nclassid<>"" then
		sql="select * from MusicList where Nclassid="+cstr(Nclassid)+" and sclassid="+cstr(sclassid)+" and classid="+cstr(classid)+"  order by id desc" 
	elseif SClassid<>"" then
		sql="select * from MusicList where SClassID="+cstr(SClassID)+" and ClassID="+cstr(Classid)+" order by id desc" 
	else
		sql="select * from MusicList where classid="+cstr(classid)+" order by id desc" 
	end if
else
	sql="select * from MusicList order by id desc"
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
	response.write "<p align='center'><b>��ʱû���ռ��κθ���<br><br><a href='javascript:history.go(-1)'>...::: �� �� �� �� :::...</a></b></p>" 
	showList
else 
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
  <tr>
    <td>
	<div align="center">
  <center>
        </center>                                                         
</div>
<div align="center">
  <center>
          <table width="766" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
            <tr> 
              <td height="1" bgcolor="#000000"></td>
            </tr>
            <tr> 
              <td bgcolor="#E7F3FF"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="28%"><img src="map/topins.gif" width="194" height="17"></td>
                    <td width="72%" valign="bottom"><a href="Artlist.asp?Name=Yxmusic&ID=13">DJҡ��</a> 
                      <img src=images/gexian.gif width=1 height=10> <a href="Artlist.asp?Name=Yxmusic&ID=12">��������</a> 
                      <img src=images/gexian.gif width=1 height=10> <a href="midi.htm">MIDI��Ʒ</a> 
                      <img src=images/gexian.gif width=1 height=10> <a href="http://www.k666.com" target="_blank">�Ҹ�����</a> 
                      <img src=images/gexian.gif width=1 height=10> <a href="f4/index.htm" target="_blank">F4����д�漯</a> 
                      <img src=images/gexian.gif width=1 height=10> <a href="http://news.kk66.com" target="_blank">��������</a> 
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr> 
              <td bgcolor="#000000" height="1" background="map/fix.gif"></td>
            </tr>
            <tr> 
              <td bgcolor="#E7F3FF"> 
                <table width="766" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="587" height="20">��<a href="index.asp"><font color="#000000">һ������ 
                      &gt;</font></a> <font color="#000000"> <a href="Albumlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>"><%=rs("Singer")%></a> 
                      &gt; <a href="Albumlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>">��ר���б�</a> 
                      &gt; �����б�</font> </td>
                    <td width="179">&nbsp;</td>
                  </tr>
                  <tr > 
                    <td colspan="2" background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </center>
</div>
<div align="center">
  <center>
          <table border="0" cellpadding="0" cellspacing="0" width="766">
            <tr>
      <td width="171" valign="top">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="170" bgcolor="#D6E3FF">
            <tr>
              <td width="100%" bgcolor="#1251E4" align="center" valign="bottom" height="18"><font color="#E7E7E7">:::</font>               
              <font color="#FFFFFF">�� �� �� ��</font> <font color="#E7E7E7">:::</font></td>
            </tr>
            <tr>
              <td align="center" height="111">  
                <div align="center">  
                  <table border="0" cellpadding="5" cellspacing="0" width="98%">  
                    <form name="search" method="post" action="search.asp"><tr>
                      <td width="100%" align="center"><input size="12" name="keyword"></td>  
                    </tr>  
                    <tr>  
                 <td width="100%" align="center"><select style="FONT-SIZE: 9pt" name="stype">  
                  <option value="Music" selected>&nbsp;��������&nbsp;</option>
                  <option value="Special">&nbsp;ר������&nbsp;</option>
                  <option value="Singer">&nbsp;��������&nbsp;</option>  
                  </select></td>
                    </tr>
                    <tr>
                      <td width="100%" align="center"><input type="submit" value="����" name="submit">&nbsp; 
                        <input type="reset" value="����" name="submit2"></td>
                    </tr></form>
                  </table>
                </div>
              </td>
            </tr>
            <tr>
             <td width="100%" bgcolor="#1251E4" align="center" valign="bottom" height="18"><font color="#E7E7E7">:::</font>                                                          
              <font color="#FFFFFF">�� �� Ŀ ¼</font> <font color="#E7E7E7">:::</font></td>
            </tr>
            <tr>
              <td align="center" bgcolor="#D6E3FF" height="410">
<div align="center">    
<center>    
<table border="0" cellpadding="5" cellspacing="0" width="98%">    
<tr>    
<td width="100%"><b>&nbsp;���˸���</b></td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=1">�����и���</a></td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=2">����Ů����</a></td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=3">�����ֶ����</a></td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=4">Ӱ�Ӻϼ�����</a></td>    
</tr>    
<tr>    
<td width="100%"><b>&nbsp;�պ�����</b></td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=5">�պ��и���</a></td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=6">�պ�Ů����</a>     
</td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=7">�պ��ֶ����</a></td>    
</tr>    
<tr>    
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=8">Ӱ�Ӻϼ�����</a></td>    
</tr>    
<tr>    
<td width="100%"><b>&nbsp;ŷ������</b></td>   
</tr>   
<tr>   
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=9">ŷ���и���</a></td>   
</tr>   
<tr>   
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=9">ŷ��Ů����</a></td>   
</tr>   
<tr>   
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=9">ŷ���ֶ����</a></td>   
</tr>   
<tr>   
<td width="100%" align="center"><a href="Artlist.asp?Name=Yxmusic&ID=9">Ӱ�Ӻϼ�����</a></td>   
</tr>   
</table>    
</center>    
</div>
             </td>
            </tr>
            <tr>
              <td align="center"></td>
            </tr>
          </table>
        </div>
      </td>
      <td width="599" valign="top" align="right">
        <table border="0" cellpadding="0" cellspacing="0" width="588">
          <tr>
            <td width="230" background="image/bg2.gif"  height="19" valign="bottom">&nbsp; 
              <b><font color="#ffffff">�� �� �� �� :::...</font></b></td>
            <td width="358" background="image/bg2.gif"  align="right" valign="bottom"></td>
          </tr>
          <tr>
            <td width="100%" colspan="2"></td>
          </tr>
          <tr>
            <td width="100%" colspan="2">
<!----------------------------Recall star Music---------------------------------->
<%
	MaxPerPage=15
	totalPut=rs.recordcount 
	if currentpage<1 then currentpage=1
	if (currentpage-1)*MaxPerPage>totalput then 
		if (totalPut mod MaxPerPage)=0 then 
			currentpage= totalPut \ MaxPerPage 
		else 
			currentpage= totalPut \ MaxPerPage + 1 
		end if 
	end if 
	if currentPage=1 then 
		showpage totalput,MaxPerPage,"songlist.asp" 
		showContent 
		showpage totalput,MaxPerPage,"songlist.asp" 
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move  (currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark 
			showpage totalput,MaxPerPage,"songlist.asp" 
			showContent 
			showpage totalput,MaxPerPage,"songlist.asp" 
		else 
			currentPage=1 
			showpage totalput,MaxPerPage,"songlist.asp" 
			showContent 
			showpage totalput,MaxPerPage,"songlist.asp" 
			showList
		end if 
	end if 
	rs.close 
end if 

sub showContent 
dim i 
i=0 
%>
<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" width="100%" bordercolor="#C0C0C0" bordercolordark="#FFFFFF">
        <form name="form" onsubmit="javascript:return lbsong();" target="lbsong" action="Yxplaylist.asp"><tr>
         
          <td width="7%" height=22 align=center bgcolor="#D6E3FF">ѡ��</td>
          <td width="47%" height=22 align=center bgcolor="#D6E3FF">��������</td>
          <td width="9%" height=22 align=center bgcolor="#D6E3FF">����</td>
          <td width="9%" height=22 align=center bgcolor="#D6E3FF">Wma����</td>
          <td width="9%" height=22 align=center bgcolor="#D6E3FF">���</td>
          <td width="9%" height=22 align=center bgcolor="#D6E3FF">���ֺ�</td>
        </tr>
<%do while not rs.eof%>
        <tr>
          <td align="center"><input type="checkbox" name="checked" value="<%=rs("id")%>"></td>
          <td> <%=i%>.<a href="#" onclick="return callpage('Yxplay.asp?id=<%=rs("id")%>');"><%=rs("MusicName")%></a></td>
          <td align=center><%=rs("hits")%></td>
          <td align=center><%if rs("Wma")<>"" then%><a href="#" onclick="return callpage('Yxplay.asp?id=<%=rs("id")%>');">����</a><%else%>����<%end if%></td>
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
  </center>
</div>
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
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
   <form method=Post action="<%=filename%>?classid=<%=classid%>&Sclassid=<%=Sclassid%>&Nclassid=<%=Nclassid%>">
    <tr>
      <td height="10"></td>
    </tr>
     <tr>
                                <td width="100%" bgcolor="#D6E3FF" align="center"> 
                                  ��<font color="#ff0000"><b><%=totalnumber%></b></font>�׸��� 
                                  <%if CurrentPage<2 then%>
                                  &nbsp;��ҳ &nbsp;��һҳ&nbsp; 
                                  <%else%>
                                  &nbsp<a href="<%=filename%>?page=1&Name=<%=classid%>&ID=<%=Sclassid%>&ArtID=<%=Nclassid%>">��ҳ</a>&nbsp; 
                                  <a href="<%=filename%>?page=<%=CurrentPage-1%>&Name=<%=classid%>&ID=<%=Sclassid%>&ArtID=<%=Nclassid%>">��һҳ</a>&nbsp; 
                                  <%
end if
if n-currentpage<1 then
%>
                                  ��һҳ &nbsp;ĩҳ 
                                  <%else%>
                                  <a href="<%=filename%>?page=<%=CurrentPage+1%>&Name=<%=classid%>&ID=<%=Sclassid%>&ArtID=<%=Nclassid%>">��һҳ</a>&nbsp; 
                                  <a href="<%=filename%>?page=<%=n%>&Name=<%=classid%>&ID=<%=Sclassid%>&ArtID=<%=Nclassid%>">ĩҳ</a>&nbsp; 
                                  <%end if%>
                                  &nbsp;ҳ��:<strong><font color="#ff0000"><%=CurrentPage%>/<%=n%></font></strong>ҳ 
                                </td>
    </tr>
    <tr>
      <td height="10"></td>
    </tr>
   </form> 
  </table>
 </center>
</div>
<% 
end function
%>
<!----------------------------Recall star Music---------------------------------->
</td>
          </tr>
          <tr>
            <td width="100%" colspan="2" height="20"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  </center>
</div>
	</td>
  </tr>
</table>
<%
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="Bottom.asp"-->