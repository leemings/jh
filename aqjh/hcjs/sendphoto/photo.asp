<!--#include file="connpic.asp"-->
<%
dim page
page=request.querystring("page")
PageSize = 9
   set buys=server.createobject("adodb.recordset")    

If Session("aqjh_grade")>=10  then
    buys.open "Select * From pic Order by id DESC",conn,3,3
else
    buys.open "Select * From pic where 批准=1 Order by id DESC",conn,3,3
end if
    buys.PageSize = PageSize
    pgnum=buys.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then buys.AbsolutePage=page
%>
<html>
<head>
<title>网友照片</title>
<style></style>
<link rel="stylesheet" href="../../css.CSS" type="text/css">
</head>
<BODY  background="../../bg.gif">
<center>
<table width=500 border=0><tr><td align=center>网 友 照 片</table>

            <table width="500" border="0" style=font-size:9pt>
              <tr> 
                <td><a href="sendphoto.asp">上传照片</a></td> 
              </tr>
            </table>
            <table border="0" cellspacing="2" cellpadding="3" width="600" align="center" >
              <% 
count=0 
do while not (buys.eof or buys.bof) and count<buys.PageSize 
%>
              <%if (count+3) mod 3 =0 then%>
              <tr> 
                <%end if%>
                <td width="10%" align=center> 
                  <table border="0" cellspacing="0" cellpadding="0" bordercolor="#006531" style=font-size:9pt>
                    <tr> 
                      <td bgcolor="#007cd0"> 
                        <div align="center"> <a href=display1.asp?id=<%=buys("id")%> target="_blank"><img border="1" width=150 height=100 src='display1.asp?id=<%=buys("id")%>'></a></td></tr>
                          <tr><td bgcolor="#007cd0" height=20 align=center><font color="#FFFFFF"><%=buys("name")%></font>　<%If Session("aqjh_grade")>=10 or  Session("aqjh_name")=buys("name") Then%>
<a href=del.asp?id=<%=buys("id")%>><font color="#FFFFFF">删除</a><%if buys("批准")=0 then%> <a href=pz.asp?id=<%=buys("id")%>><font color="#FFFFFF">批准</a><%end if
end if%></div>
                      </td>
                    </tr>
                  </table>
                </td>
                <%if (count+4) mod 3 =0 then%>
              </tr><tr><td height=24>

</td></tr>
              <%end if%>
              <%
buys.movenext
count=count+1
loop%>
            </table>
            <table border="0" cellspacing="0" cellpadding="2" width="500" align="center" style=font-size:9pt>
              <tr> 
                <td align=left> 
                  <div align="right">[<a href="photo.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a href="photo.asp?page=<%=page+1%>">下一页</a>]</div>
                </td>
              </tr>
            </table>
</body>
</html>
<%
function musers()   
dim tmprs   
    tmprs=conn.execute("Select count(*) As lar_id from larchives")   
    musers=tmprs("name")   
	set tmprs=nothing   
	if isnull(musers) then musers=0   
end function   
%>