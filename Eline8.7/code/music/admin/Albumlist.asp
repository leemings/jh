<!--#include file="conn.asp"-->
<!--#include file="Star.INC"-->
<%
classid=request.QueryString("Name")
Sclassid=request.QueryString("ID")
Nclassid=request.QueryString("ArtID")
id=request.QueryString("Name")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

id=request.QueryString("id")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

id=request.QueryString("ArtID")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
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

	response.write "<p align='center'><br>Sorry, �� ʱ û �� �� �� ר ��<br><br><a href=javascript:history.go(-1)><br> �� �� �� </a><br><br></p>" 
else 
    MaxSpecialList=30	
i=0
%>
<head>
<title>"<%=rs("Nclass")%>"ר���б�::::::{���������� }::::::http://WWW.119MUSIC.COM/</title>
</head>
<!--#include file="Top.asp"-->
<table width="766" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="20" bgcolor="#F79618">��<a href="index.asp"><font color="#000000">���������� 
      &gt;</font></a> <font color="#000000"><a href="Artlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>"><%=rs("SClass")%> &gt;</a> <%=rs("Nclass")%> &gt; ר���б� 
    <SCRIPT src="119.js"></SCRIPT></font></td>
  </tr>
  <tr > 
    <td background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
  </tr>
</table>
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
                <td width="19%" align="center" height="120"> 
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
                <td width="47%">ר�����ƣ�<a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>">{<%=rs("name")%>}</a><br>
                  <br>
                  �������ڣ�<%=rs("Times")%></td>
                <td width="34%">���������<%=rs("hits")%><br>
                  <br>
                  <a href="MusicList.asp?AlbumID=<%=rs("Specialid")%>">��������</a></td>
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
    <p>��</p></td>
  </tr>
</table>
<%
end if
set Trs=nothing
set rs=nothing
conn.close
set conn=nothing
%><!--#include file="Bottom.asp"-->