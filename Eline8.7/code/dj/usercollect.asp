
<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<!--#include file="jhconst.asp"-->

<%
action=request("action")

if action="add" then
username=session("sjjh_name")
set rs=server.createobject("adodb.recordset")
sql="select * from user where username='"&session("sjjh_name")&"'"
rs.open sql,conn,1,3
if rs.eof then
	response.write "�Բ��𣡴˹�����ʱ���Σ�"
	Response.end
else
		 end if
	rs.close
	set rs=nothing
	Musicid=request.QueryString("songid")
	set rs=conn.execute("SELECT * FROM MusicDJ where id="+Musicid)
	MusicName=rs("MusicName")
	MusicType=rs("MusicType")
	rs.close
	set rs=nothing
	set rs=server.createobject("adodb.recordset")
	Musicid=request.QueryString("songid")
		sql="select * from collection where UserName='"&session("sjjh_name")&"'"
	rs.open sql,conn,1,3
		 if RS.recordcount>Collect then
		   response.write "<P align=center><font color=red>�� �� ��</font><br><br>�Բ������������ղؼ�������<br><br><a href=UserCollect.asp?action=show>������</a>" 
		   response.end
		  end if
	rs.close
	sql="select * from collection where UserName='"&session("sjjh_name")&"' and Musicid="&Musicid
	rs.open sql,conn,1,3
	if not rs.EOF then
		errmsg=errmsg+"<br><li>���Ѿ��ղ��˴������ˣ�</li>"
		call error()
		response.end
	else

		rs.AddNew
		rs("UserName")=session("sjjh_name")
		rs("Musicid")=Musicid
		rs("MusicName")=MusicName
		rs("MusicType")=MusicType
		rs.Update
	end if
	rs.Close
	Response.Redirect "UserCollect.asp?action=show"

elseif action="del" then
	id=request("id")
	set rs=conn.execute("delete FROM collection where UserName='"&session("sjjh_name")&"' and id="&id)
	Response.Redirect "UserCollect.asp?action=show"

elseif action="show" then
%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ҵ������ղؼ�</title>
</head>
<style>
td
{
FONT-WEIGHT: normal; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; color: #000000
}
A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #000000
}
A:visited {
	COLOR: #000000
}
A:active {
	COLOR: #000000
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline overline
}

input
{
		background-image: url('images/inbg.gif');border:1px solid #CE9A00; padding-left: 0
}
</style>
<script src=JS/Js.js></script>
<body topmargin="0" leftmargin="0">
      <div align="center">
        <center>
      <table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#000000">
<form name=form2 onsubmit='javascript:return DJplay_song();' target=HeiRui_Studio_Player action="djplay.asp">
        <tr>
          <td align=center colspan=4 bgcolor="#FFCC66"><b><font color=#cc0000><%=membername%></font>�������ղؼ�
          <input type=button value=�Ŵ󴰿� onclick=window.resizeTo(420,700)>
          <input type=button value=��С���� onclick=window.resizeTo(420,300)>
          <input type=button value=ˢ�� onclick=history.go(-0) ></b></td>
        </tr>
        <tr>
          <td align=center bgcolor="#FFB720" width="5%">ID</td>
          <td align=center bgcolor="#FFB720">��������</td>
          <td align=center bgcolor="#FFB720">����</td>
          <td align=center bgcolor="#FFB720">ɾ��</td>
        </tr>

<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from collection where UserName='"&session("sjjh_name")&"'"
	rs.open sql,conn,1,1
	if rs.EOF and rs.BOF then
%>
        <tr><td align=center colspan=4>����δ�ղ��κ�����</td></tr>
<%
	else
		do while not rs.EOF 
%>
        <tr>
          <td align=center bgcolor="#FFCC66">
    <input type=checkbox name=songid value=<%=rs("Musicid")%> style="width:15;height:15"></td>
          <td align=center bgcolor="#FFCC66"><%=rs("MusicName")%>��</td>
           <td align=center bgcolor="#FFCC66">
<img src=images/radio2.gif alt=���������� style="cursor:hand" onclick="window.open('<%if rs("musictype")="RealPlayer" then%>djplay.asp<%elseif rs("musicType")="MediaPlayer" then%>djplay_media.asp<%else%>djplay_swf.asp<%end if%>?songid=<%=rs("Musicid")%>','HeiRui_Studio_Player','width=276,height=343');"></td>
          <td align=center bgcolor="#FFCC66">
<a onclick=history.go(-0) href="UserCollect.asp?action=del&id=<%=rs("id")%>">
<img border="0" src="images/del.gif">ɾ��</a></td>
        </tr>
<%
		rs.MoveNext
		loop
%>
        <tr>
          <td align=center bgcolor="#FFB720" colspan="4">
		<input type=button name=chkall value='ȫ ѡ' onclick='CheckAll(this.form)' title=ѡ����ʾ����������>
       	<input type=button name=chkOthers value='�� ѡ' onclick='CheckOthers(this.form)' title=����ѡ������>
		<input type=submit value='���벥��' title=����ѡ�����������������ٵ������> </td>
          </tr>
          </form>
      </table>
        </center>
      </div>

</body>
</html>
<%
	end if
	rs.Close

else
	errmsg=errmsg+"������������ϵ����Ա��"
	founderr=true
end if
set rs=nothing
conn.close
set conn=nothing
if founderr=true then
	call error()
end if
%>