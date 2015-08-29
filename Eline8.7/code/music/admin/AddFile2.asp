<%PageName="1-1sort"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<%
Classid=Request.QueryString ("Classid")
SClassid=Request.QueryString ("SClassid")
if Classid="" and SClassid="" then
	AskClassid=Request.QueryString ("AskClassid")
	AskSClassid=Request.QueryString ("AskSClassid")
	Page=Request.QueryString ("")
end if
set rs=server.createobject("adodb.recordset")
set rs2=server.createobject("adodb.recordset")
set rs3=server.createobject("adodb.recordset")
%>
     <div align="center">
      <center>
      <table border="1" width="90%" cellspacing="0" cellpadding="0" class="TableLine" bordercolorlight="#CC0066">
        <tr>
          <td width="100%" height="20" bgcolor="#CC0066" align=center><a href="Addfile3.asp?classid=1"><font color="white"><b>添 加 编 辑 专 辑 (第二步)&nbsp;&nbsp;&nbsp;点 这 里 可 直 接 添 加 > > ></b></font></a></td>
        </tr>
        <tr>
        <td height="20" align="center"><a href="#a"><B>A</B></a>  
            <a href="#b"><B>B</B></a>  <a href="#c"><B>C</B></a> 
             <a href="#d"><B>D</B></a>  <a href="#e"><B>E</B></a> 
             <a href="#f"><B>F</B></a>  <a href="#g"><B>G</B></a> 
             <a href="#h"><B>H</B></a>  <a href="#i"><B>I</B></a> 
             <a href="#j"><B>J</B></a>  <a href="#k"><B>K</B></a> 
             <a href="#l"><B>L</B></a>  <a href="#m"><B>M</B></a> 
             <a href="#n"><B>N</B></a>  <a href="#o"><B>O</B></a> 
             <a href="#p"><B>P</B></a>  <a href="#q"><B>Q</B></a> 
             <a href="#r"><B>R</B></a>  <a href="#s"><B>S</B></a> 
             <a href="#t"><B>T</B></a>  <a href="#u"><B>U</B></a> 
             <a href="#v"><B>V</B></a>  <a href="#w"><B>W</B></a> 
             <a href="#x"><B>X</B></a>  <a href="#y"><B>Y</B></a> 
             <a href="#z"><B>Z</B></a></td>
         </tr>
<%
sql="select * from class order by classid"
rs.open sql,conn,1,1
if rs.eof then
%>
<%
else
	do while not rs.eof
%>
<%	rs.movenext
	loop
end if
rs.close
%>
<%
		sql3="select * from Nclass where SClassid="&SClassid&" order by Abcd"
		rs3.open sql3,conn,1,1
		MaxPerPage=10000

	sql="select * from class where Classid="&Classid
	rs.open sql,conn,1,1
%>
<%	if not rs.eof then%>
<%	end if%>
<%
	if rs3.eof then
%>
              <tr><td>尚无任何分类</td></tr>
<%
	else
		k=0
		do while not rs3.eof
			k=k+1
							if thischar<>rs3("Abcd") then
								thischar=rs3("Abcd")
								if i=0 then
								 Response.Write "<tr><td width=50% bgcolor=#CC0066>&nbsp;&nbsp;<a name="&thischar&"><b><font color=white face=Verdana, Arial, Helvetica, sans-serif size=4>"&thischar&"</font></b></a></td></tr>"
								 else
								 Response.Write "<tr><td width=50% bgcolor=#CC0066>&nbsp;&nbsp;<a name="&thischar&"><b><font color=white face=Verdana, Arial, Helvetica, sans-serif size=4>"&thischar&"</font></b></a></td></tr>"
								 end if
end if
%>
		<tr>
          <td width="100%" height="25" align=left>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                   <form method="POST" action="NClassSave.asp?act=rename&NClassid=<%=rs3("NClassid")%>" id=Nform<%=k%> name=Nform<%=k%>>
                    <tr>
                      <td width="20%">&nbsp;&nbsp;<a href='AddFile3.asp?Classid=<%=rs3("Classid")%>&SClassid=<%=rs3("SClassid")%>&NClassid=<%=rs3("NClassID")%>'><%=rs3("NClass")%></a></td>
                      <td width="40%"><a href='AddFile3.asp?Classid=<%=rs3("Classid")%>&SClassid=<%=rs3("SClassid")%>&NClassid=<%=rs3("NClassID")%>'><---进入添加专辑</a></td>
                      <td width="40%"><a href=AddFileList.asp?Classid=<%=rs3("Classid")%>&SClassid=<%=rs3("SClassid")%>&NClassid=<%=rs3("NClassid")%>>浏览该歌手的所有专辑</a></td>
                    </tr>
                   </form>
                  </table>
                </div>
		  </td>
        </tr>
<%
			if k>=MaxPerPage then exit do
		rs3.movenext
		loop
	end if
	rs3.close
%>

      </table>
     </center>
    </div>
<% 
set rs=nothing
set rs3=nothing
conn.close
set conn=nothing
%></body></html>


