<!--#include file="conn.asp"-->
<%
act=request("act")
	select case act
	case "加入歌曲列表"
		call add()
	case "删除"
		call del()
	case "清空列表"
		call clear()
	case "none"
		call none()
	case else
	end select

%>

<%
sub add()
%>
<SCRIPT language="javascript">
 parent.form2.playlist.options.length = 0;
 parent.form2.playlist.options[parent.form2.playlist.options.length]= new Option("---==播放列表==---","");
</SCRIPT>
<%
    set rs=server.createobject("adodb.recordset")
    if request.cookies("playlist")<>"" then
    response.cookies("playlist")=request.cookies("playlist")+","+replace(request("songid")," ","")
    else
    response.cookies("playlist")=replace(request("songid")," ","")
    end if
    id=request.cookies("playlist")
    sql="select * from MusicDJ where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
	MusicName=rs("musicname")
	id=rs("id")
%><script>parent.document.all.playlist.options.add(new Option('<%=musicname%>','<%=id%>'))</script>
<%
rs.movenext
wend
rs.Close
end sub
%>

<%
sub del()
%>
<SCRIPT language="javascript">
 parent.form2.playlist.options.length = 0;
 parent.form2.playlist.options[parent.form2.playlist.options.length]= new Option("---==播放列表==---","");
</SCRIPT>
<%
response.cookies("playlist")=replace(request.cookies("playlist"),","&request("playlist"),"")
response.cookies("playlist")=replace(request.cookies("playlist"),request("playlist")&",","")
response.cookies("playlist")=replace(request.cookies("playlist"),request("playlist"),"")
    set rs=server.createobject("adodb.recordset")
    if request.cookies("playlist")<>"" then
    id=request.cookies("playlist")
    sql="select * from MusicDJ where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
	MusicName=rs("musicname")
	id=rs("id")
%><script>parent.document.all.playlist.options.add(new Option('<%=musicname%>','<%=id%>'))</script>
<%
rs.movenext
wend
rs.Close
end if
end sub
%>


<%
sub clear()
%>
<SCRIPT language="javascript">
 parent.form2.playlist.options.length = 0;
 parent.form2.playlist.options[parent.form2.playlist.options.length]= new Option("---==播放列表==---","");
</SCRIPT>
<%
response.cookies("playlist")=""
end sub
%>



<%
sub none()
    set rs=server.createobject("adodb.recordset")
    if request.cookies("playlist")<>"" then
    id=request.cookies("playlist")
    sql="select * from MusicDJ where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
	MusicName=rs("musicname")
	id=rs("id")
%><script>parent.document.all.playlist.options.add(new Option('<%=musicname%>','<%=id%>'))</script>
<%
rs.movenext
wend
rs.Close
end if
end sub
%>