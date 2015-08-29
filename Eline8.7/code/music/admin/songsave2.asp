<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<!--#include file="Star.inc"-->
<%
page=trim(request.querystring("page"))
AskClassid=trim(request.querystring("AskClassid"))
AskSClassid=trim(request.querystring("AskSClassid"))
AskNClassid=trim(request.querystring("AskNClassid"))
Wma=trim(request.form("Wma"))
MusicName=trim(request.form("MusicName"))
classid=trim(request.form("classid"))
Sclassid=trim(request.form("Sclassid"))
Nclassid=trim(request.form("Nclassid"))
Specialid=trim(request.form("Specialid"))

founerr=false
if act<>"SetIsGood" then
act=request("act")
end if

if founderr=true then
	call error()
else
	set rs=server.createobject("adodb.recordset")
	if act<>"SetIsGood" and act<>"SetIsChecked" then
		sql="select NClass from NClass where NClassid="&Nclassid
		rs.open sql,conn,1,1
		singer=rs("Nclass")
		rs.close
	end if

	if act="edit" and request("id")<>"" then
		sql="select * from MusicList where id="&request("id")
		rs.open sql,conn,1,3
		rs("Wma")=trim(Wma)
		rs("MusicName")=MusicName
		rs("classid")=classid
		rs("Sclassid")=Sclassid
		rs("Nclassid")=Nclassid
		rs("Specialid")=Specialid
		rs("singer")=trim(singer)
		rs.update
		rs.close

		sql="select * from Box where Musicid="&request("id")
		rs.open sql,conn,1,3
		if not rs.eof then
			rs("MusicName")=MusicName
			rs.update
		end if
		rs.close

		call Success
		Response.End 
	elseif act="add" then
		sql="select * from MusicList where (id is null)" 
		rs.open sql,conn,1,3
		rs.addnew
	    rs("Wma")=Wma
		rs("MusicName")=MusicName
		rs("classid")=classid
		rs("Sclassid")=Sclassid
		rs("Nclassid")=Nclassid
		if Specialid<>"" then rs("Specialid")=Specialid
		rs("singer")=singer
		rs.update
		rs.close
		call Success
		Response.End 
	elseif act="SetIsGood" then
		sql="select IsGood from MusicList where id="&request("id") 
		rs.open sql,conn,1,3
		if not rs.EOF then
			if rs("IsGood")=true then
				rs("IsGood")=false
			else
				rs("IsGood")=true
			end if
			rs.update
		end if
		rs.close
	else
		errmsg=errmsg+"<li>操作错误！请联系管理员</li>"
		call error()
		Response.End
	end if
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "Songlist.asp?Classid="+AskClassid+"&SClassid="+AskSClassid+"&NClassid="+AskNClassid+"&Page="+Page
end if
sub Success
%>
<%PageName="SongSave"%>
<!--#include file="top.asp"-->
<div align="center">
<center>
<table border="0" width="80%" cellspacing="1" cellpadding="1">
  <tr>
    <td>　
      <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <tr align="center">
        <td width="100%" height="20" colspan=2 bgcolor="#FF1171" align=center><font color="white"><b>歌曲<%if act="add" then%>添加<%else%>修改<%end if%>成功</b></font></td>
        </tr>
        <tr>
          <td>
           <table border="1" width="100%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
             <tr>
              <td width="25%" align="right">歌曲：</td>
              <td width="75%"><%=MusicName%>　</td>
            </tr>
            <tr>
              <td align="right">歌手：</td>
              <td><%=Singer%>　</td>
            </tr>
            <tr>
              <td align="right">试听地址：</td>
              <td><%=Wma%>　</td>
            </tr>
            <tr>
              <td colspan="2" height="15" align=center><input type="button" name="button2" value=" 继续<%if act="add" then%>添加<%else%>修改<%end if%> " onclick="javascript:history.go(-1)">&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button1" value=" 返回 " onclick="javascript:history.go(-2)"></td>
            </tr>
          </table>
          </td>
        </tr>
      </table>
 　 </td>
  </tr>
</table>
</div>
</body>
</html>
<%end sub%>



