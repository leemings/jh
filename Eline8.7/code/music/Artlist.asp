<!--#include file="conn.asp"-->
<!--#include file="Star.INC"-->
<%
Classid=request.QueryString("Name")
Sclassid=request.QueryString("ID")
id=request.QueryString("Name")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if
id=request.QueryString("ID")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

set rs=server.createobject("adodb.recordset")
%>
<!--#include file="Top.asp"-->
<div align="center"> 
  <table width="766" border="0" cellpadding="0" cellspacing="0" bgcolor="#F79618">
    <tr> 
      <td width="587" height="20" bgcolor="#66CCFF">　<a href="index.asp"><font color="#000000">一线视听
        &gt;</font></a> <font color="#000000">歌手列表</font></td>
      <td width="179" bgcolor="#66CCFF">　</td>
    </tr>
    <tr > 
      <td colspan="2" background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
    </tr>
  </table>
</div>
<div align="center"> 
  <table width="766" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
    <tr> 
      <td width="129" valign="top" bgcolor="#E7F3FF"> 
        <!--#include file="leftmenu.asp"-->
      </td>
      <td width="1" bgcolor="#FF0000" background="map/fix2.gif"></td>
      <td width="750"> <table width="100%" border="0">
          <tr bordercolorlight="#CC0066" class="TableLine">
        </table>
        <table border="0" cellpadding="0" cellspacing="0" width="100%" align="center" bgcolor="#F0F8FF">
          <tr> 
            <td width="100%" colspan="5" bgcolor="#F0F8FF"></td>            
<%
if Sclassid<>"" then
set rs=conn.execute("SELECT * FROM Nclass where Sclassid="&Sclassid&" order by Abcd")
if not rs.eof then
i=0
do while not rs.eof
i=i+1
if thischar<>rs("Abcd") then
thischar=rs("Abcd")
if i=1 then
Response.Write "</tr><tr><td width=100% bgcolor=#B4DEF8 height=25 align=center colspan=5><b><font face=Verdana, Arial, Helvetica, sans-serif size=4>"&thischar&"</font></b></td></tr><tr><td height=1 colspan=5 background=map/fix.gif></tr><tr>"
else
Response.Write "</tr><tr><td height=1 colspan=5 background=map/fix.gif></tr><tr><td width=100% height=25 bgcolor=#B4DEF8 align=center colspan=5><b><font face=Verdana, Arial, Helvetica, sans-serif size=4>"&thischar&"</font></b></td></tr><tr><td height=1 colspan=5 background=map/fix.gif></tr><tr>"
end if
end if
%>
            <td width=33% height=22 align=center bgcolor=#F0F8FF>&nbsp;&nbsp;<a href="Albumlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>"><%=rs("NClass")%></a></td>
            <%
if (i mod 3)=0 and i>=3 then
%>
          </tr>
          <tr> 
            <%
end if
rs.movenext
loop
else
%>
            <%
end if
rs.close
%>
          </tr>
        </table>
        <br>
      </td>
    </tr>
  </table>
</div>
<%
end if
set rs=nothing
conn.close
set conn=nothing
%><!--#include file="Bottom.asp"-->