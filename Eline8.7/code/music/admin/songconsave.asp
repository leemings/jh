<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from MusicList"
rs.open sql,conn,1,3
Num=100
for i=1 to Num
	Classic=request.form("Classic"&i)
	ConClassic=split(Classic,",")
	MusicName=trim(request.form("MusicName"&i))
	Wma=trim(request.form("Wma"&i))
    url=request.form("url")

	if MusicName="" or isnull(MusicName) or Wma="" or isnull(Wma) then
		Num=i-1
		exit for
	end if
	
	set NRs=server.createobject("adodb.recordset")
	Nsql="select NClass from NClass where NClassid="&ConClassic(2)
	NRs.open Nsql,conn,1,1
	if not NRs.eof then
		NClass=NRs("NClass")
	end if
	NRs.close
	
	rs.addnew
	rs("MusicName")=MusicName
	rs("Singer")=NClass
	rs("Classid")=ConClassic(0)
	rs("SClassid")=ConClassic(1)
	rs("NClassid")=ConClassic(2)
	rs("Specialid")=ConClassic(3)
	rs("Wma")=url&Wma
	rs.update
next
rs.close
call Success

sub Success
%>
<!--#include file="top.asp"-->
<BR><BR><BR><BR><BR>
<div align="center">
<center>
<table border="0" width="60%" cellspacing="1" cellpadding="1">
  <tr>
    <td>　
      <table border="1" width="100%" cellspacing="0" cellpadding="5" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <tr align="center">
         <td width="100%" height="20" colspan=2 bgcolor="#FF1171" align=center><font color="white"><b>歌曲批量添加成功<br>共添加了 <%=Num%> 首歌曲</b> </td>
        </tr>
        <tr>
          <td align=center><input type="button" name="button1" value="返回" onclick="javascript:history.go(-2)">&nbsp;&nbsp;<input type="button" name="button2" value="继续添加" onclick="javascript:history.go(-1)"></td>
        </tr>
      </table>
 　 </td>
  </tr>
</table>
</div>
</body>
</html>
<%end sub
set rs=nothing
conn.close
set conn=nothing
%>


