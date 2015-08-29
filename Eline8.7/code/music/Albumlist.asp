<!--#include file="conn.asp"-->
<!--#include file="Star.INC"-->
<%
classid=request.QueryString("Name")
Sclassid=request.QueryString("ID")
Nclassid=request.QueryString("ArtID")
id=request.QueryString("Name")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

id=request.QueryString("id")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

id=request.QueryString("ArtID")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

%>

<%
set rs=server.createobject("adodb.recordset")

if classid<>"" then
	if Nclassid<>"" then
		sql="SELECT Specialid,name,Times,hits,pic ,SClassid,SClass ,Nclass FROM Special where Nclassid="+cstr(Nclassid)+" order by Specialid desc"
	else
		sql="SELECT Specialid,name,Times,hits,pic ,SClassid,SClass,Nclass FROM Special where classid="+cstr(classid)+" order by Specialid desc"
	end if
else
		sql="SELECT Specialid,name,Times,hits,pic  ,SClassid,SClas ,Nclass FROM Special order by Specialid desc"
end if
rs.open sql,conn,1,1

if rs.eof and rs.bof then 

	response.write "<p align='center'><br>Sorry, 暂 时 没 有 任 何 专 辑<br><br><a href=javascript:history.go(-1)><br> 请 返 回 </a><br><br></p>" 
else 
    MaxSpecialList=30	
i=0
%><head>
<title>"<%=rs("Nclass")%>"专辑列表</title>
</head>
<!--#include file="top.asp"-->
<table width="766" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="20" bgcolor="#66CCFF">　<a href="index.asp"><font color="#000000">一线视听 
      &gt;</font></a> <font color="#000000"><a href="Artlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>"><%=rs("SClass")%> 
      &gt;</a> <%=rs("Nclass")%> &gt; 专辑列表</font></td>
  </tr>
  <tr > 
    <td background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
  </tr>
</table>
<table width="764" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
  <tr>
    <td width="160" valign="top" bgcolor="#E7F3FF"><!--#include file="leftmenu1.asp" --></td>
    <td width="1" bgcolor="#FF0000" background="map/fix2.gif"></td>
    <td width="603"> <br>
      <table border="0" cellspacing="1" cellpadding="0" align="center" width="99%">
        <tr> 
          <td height="22" background="images/menu_bg.gif">　　<font color="#FFFFFF"><b>专 
            辑 列 表</b></font></td>
        </tr>
        <tr> 
          <td bgcolor="#E4F1FA" height="22"> 
            <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
              <%
do while not rs.eof
i=i+1
%>
              <tr> 
                <td width="18%" align="center" height="120"> <table width="95" border="0" cellspacing="1" cellpadding="0" height="95">
                    <tr> 
                      <td background="map/bgkuan.gif"> <table width="80" border="0" cellspacing="0" cellpadding="0" align="center" height="80">
                          <tr> 
                            <td><a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>"><img height="80" width="80" src="<%=rs("pic")%>" border="0"></a></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
                <td width="82%"><TABLE style="BORDER-COLLAPSE: collapse" 
                        borderColor=#ff6600 cellSpacing=0 cellPadding=4 
                        width="100%" border=1>
                    <TBODY>
                      <TR> 
                        <TD width="19%" bgColor=#f7ebe7><FONT 
                              color=#ff6600>专集名称：</FONT></TD>
                        <TD width="44%"><a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>"><%=rs("name")%></a></TD>
                        <TD width="18%" bgColor=#f7ebe7><FONT 
                              color=#ff6600>歌手类别：</FONT></TD>
                        <TD width="19%"><font color="#000000"><a href="Artlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>"><%=rs("SClass")%></a></font></TD>
                      </TR>
                      <TR> 
                        <TD width="19%" bgColor=#f7ebe7><FONT 
                              color=#ff6600>发行日期：</FONT></TD>
                        <TD width="44%"><%=rs("Times")%></TD>
                        <TD width="18%" bgColor=#f7ebe7><FONT 
                              color=#ff6600>点击次数：</FONT></TD>
                        <TD width="19%"><%=rs("hits")%></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top width="19%" bgColor=#ffffff><FONT 
                              color=#ff6600>专辑试听：</FONT></TD>
                        <TD width="81%" bgColor=#ffffff 
                              colSpan=3><a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>">进入试听</a></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
                <%
if (i mod (MaxSpecialList/30)=0) and i>=(MaxSpecialList/30) then
%>
              </tr>
              <%
end if
if i>=MaxSpecialList then exit do
rs.movenext
loop
%>
            </table>
          </td>
        </tr>
      </table>
      <p>　</p>
      
    </td>
  </tr>
</table>
<%
end if
set Trs=nothing
set rs=nothing
conn.close
set conn=nothing
%><!--#include file="Bottom.asp"-->