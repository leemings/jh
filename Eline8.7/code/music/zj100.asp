<!--#include file="conn.asp"-->
<!--#include file="Star.INC"-->
<!--#include file="Top.asp"-->

<table width="766" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E7F3FF">
  <tr bgcolor="#66CCFF"> 
    <td width="587" height="20">��<a href="index.asp"><font color="#000000">һ������</font></a> 
      <font color="#000000"> ����ר��</font></td>
    <td width="179">��</td>
  </tr>
  <tr > 
    <td colspan="2" background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
  </tr>
</table>
<%
set rs=server.createobject("adodb.recordset")

set rs=conn.execute("SELECT top 50 Specialid,name,nclass,pic,Times,hits FROM Special order by hits desc") 

if not Rs.eof then
%>        
<table width="766" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
  <tr>
    <td width="129" valign="top" bgcolor="#E7F3FF"><!--#include file="leftmenu.asp"--></td>
    <td width="1" bgcolor="#FF0000" background="map/fix2.gif"></td>
    <td width="750"> <br>
      <table border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000" width="90%">
        <tr> 
          <td bgcolor="#5AA6DE" height="22">����<font color="#FFFFFF"><b><font color="#000000">ר 
            �� �� ��</font></b></font></td>
        </tr>
        <tr> 
          <td bgcolor="#E4F1FA" height="22"> 
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <%
do while not rs.eof
i=i+1
%>
              <tr> 
                <td width="24%" align="center" height="120"> 
                  <table width="95" border="0" cellspacing="1" cellpadding="0" height="95">
                    <tr> 
                      <td background="map/bgkuan.gif"> 
                        <table width="80" border="0" cellspacing="0" cellpadding="0" align="center" height="80">
                          <tr> 
                            <td><a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>"><img height="80" width="80" src="<%=rs("pic")%>" border="0"></a></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
                <td width="48%">ר�����ƣ�<a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>">[<%=rs("name")%>]</a><br>
                  <br>
                  �������ڣ�<%=rs("Times")%></td>
                <td width="28%">���������<font color="#FF0000"><%=rs("hits")%></font><br>
                  <br>
                  <a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>">��������</a></td>
                             </tr>
         
                    <%
					if m>=50 then exit do
					rs.movenext 
					loop
					else
					response.write "������¼"
					end if
					rs.close
					%>
            </table>
          </td>
        </tr>
      </table>
    <p>��</p></td>
  </tr>
</table>
<!--#include file="Bottom.asp"-->