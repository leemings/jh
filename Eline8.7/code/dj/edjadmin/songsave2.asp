<!--#include file="conn.asp"-->
<!--#include file="../const.asp"-->
<!--#include file="checkadmin.inc"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from MusicDJ"
rs.open sql,conn,1,3
addnum=request.form("addnum")
for i=1 to addnum
	Num=Num+1
	SClassid=trim(request.form("SClassid"&i))
	MusicName=trim(request.form("MusicName"&i))
	LF_Path=trim(request.form("LF_Path"&i))
	MusicType=trim(request.form("MusicType"&i))
	DJUser=request.form("DJUser"&i)
	if MusicName="" then
		Num=i-1
		exit for
	end if
	
	rs.addnew
	rs("MusicName")=replace(MusicName,"'","’")
	rs("SClassid")=sclassid
	rs("LF_Path")=LF_Path
	rs("DateAndTime")=now()
	rs("MusicType")=MusicType
	rs.update
next
rs.close
set rs=nothing

call Success

sub Success
%>

      <div align="center">
        <center>

      <table border="1" width="45%" cellspacing="1" bordercolor="#666666">
        <tr align="center">
          <td width=100% align=center height=20 >舞曲批量添加成功<br>共添加了<font color=""><%=Num%>首歌曲</td>
        </tr>
        <tr>
          <td>
          <p align="center">
          <input type="button" name="button1" value="返回" onclick="javascript:history.go(-2)">&nbsp;<input type="button" name="button2" value="继续添加" onclick="javascript:history.go(-1)"></td>
        </tr>
      </table>
         
        </center></div>

         
<%end sub%>