<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����ʹ�����ֺ����Ƚ��뽭�������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT ����,��� FROM �û� WHERE ����='"& session("sjjh_name") &"'",conn,1,3
if rs("����")<1000000 or rs("���")<5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "����û1000000�������Ҳ���5��������ʹ�����ֺУ�"
	response.end
end if
rs("����")=rs("����")-1000000
rs("���")=rs("���")-5
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%PageName="Box"%>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<%
action=request("action")
set rs=server.createobject("adodb.recordset")
if action="add" then
	Musicid=request.QueryString("id")
	sql="select * from Box where UserName='"&session("sjjh_name")&"' and Musicid="&Musicid
	rs.open sql,conn,1,3
	if not rs.EOF then
		errmsg=errmsg+"<SCRIPT language=JavaScript>alert(' ���Ѿ��ղ��˴˸�����! ');javascript:top.window.close();</SCRIPT>"
		call error()
		response.end
	else
		set rs2=server.createobject("adodb.recordset")
		sql2="select * from MusicList where id="&Musicid
		rs2.open sql2,conn,1,1
		if (rs2.EOF and rs2.BOF) then
			errmsg=errmsg+"<SCRIPT language=JavaScript>alert(' ����ȷѡ�����! ');javascript:top.window.close();</SCRIPT>"
			call error()
			response.end
		else
			Wma=rs2("Wma")
			MusicName=rs2("MusicName")
			Singer=rs2("Singer")
			Classid=rs2("Classid")
			SClassid=rs2("SClassid")
			NClassid=rs2("NClassid")
		end if
		rs2.close
		set rs2=nothing
		rs.AddNew
		rs("UserName")=session("sjjh_name")
		rs("Musicid")=Musicid
		rs("MusicName")=MusicName
		rs("Singer")=Singer
		rs("Wma")=Wma
		if Wma<>"" then
			rs("Wma")=true
		end if
		rs("Classid")=Classid
		rs("SClassid")=SClassid
		rs("NClassid")=NClassid
		rs.Update
	end if
	rs.Close
	Response.Redirect "Box.asp?action=show"
elseif action="del" then
	id=request("id")
	set rs=conn.execute("delete FROM Box where UserName='"&session("sjjh_name")&"' and id="&id)
	Response.Redirect "Box.asp?action=show"
elseif action="show" then
%>
<!--#include file="top.asp"-->

  <table width=766 border=0 align=center cellpadding=0 cellspacing=0 bgcolor="ffffff">
    <td align="center"><br>
      <table border="1" width="96%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#56B0F4" bordercolordark="#FFFFFF">
        <tr>
          <td bgcolor=#FAFAFA align=center colspan=8><FONT color="red"><b><%=session("sjjh_name")%></B></FONT> �����ֺ� [Ϊ���ǽ�Լ�ռ�,���������������ֺ�.лл!]</td>
        </tr>
        <form name="form" onsubmit="javascript:return lbsong();" target="lbsong" action="Yxplaylist.asp">
          <td width="6%" height=22 align=center bgcolor="#B4DEF8">ѡ��</td>
          <td width="50%" height=22 align=center bgcolor="#B4DEF8">��������</td>
          <td width="20%" height=22 align=center bgcolor="#B4DEF8">����</td>
          <td width="8%" height=22 align=center bgcolor="#B4DEF8">����</td>
          <td width="8%" height=22 align=center bgcolor="#B4DEF8">���</td>
          <td width="8%" height=22 align=center bgcolor="#B4DEF8">ɾ��</td>
        </tr>
<%
	sql="select * from Box where UserName='"&session("sjjh_name")&"'"
	rs.open sql,conn,1,1
	if rs.EOF and rs.BOF then
%>
        <tr><td align=center colspan=8>����δ�ղ��κθ���</td></tr>
<%
	else
    do while not rs.EOF 
	i=i+1
%>
        <tr>
          <td align="center"><input type="checkbox" name="checked" value="<%=rs("Musicid")%>"></td>
          <td> <%=i%>.<a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','tt90','scrollbars=no,resizable=no,width=518,height=355,menubar=no,top=168,left=168')" ><%=rs("MusicName")%></a></td>
          <td align=center><a href="Albumlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>"><%=rs("Singer")%></a></td>
          <td align=center><a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','tt90','scrollbars=no,resizable=no,width=518,height=355,menubar=no,top=168,left=168')" >����</a></td>
          <td align=center><a href="#" onClick="window.open('Mailsong.asp?id=<%=rs("id")%>','','scrollbars=no,resizable=no,width=419,height=151,menubar=no,top=98,left=198')">����</a></td>
          <td align=center><a href="Box.asp?action=del&id=<%=rs("id")%>">ɾ��</a></td>
        </tr>
<%
	rs.MoveNext
    loop
%>
    <tr>
   <td align="center" colspan="6" height="22">
   <input type="button" name="chkall" value="ȫѡ" onclick="CheckAll(this.form)" title="ѡ����ʾ�����и���">&nbsp;
   <input type="button" name="chkOthers" value="��ѡ" onclick="CheckOthers(this.form)" title="����ѡ�����">&nbsp;
   <input type="submit" name="B1" value="����" title="����ѡ���������ĸ������ٵ������">
   </td>
   </tr>
   </form>
 </table>
         
    </table> 

<!--#include file="Bottom.asp"-->
<%
	end if
	rs.Close
else
	errmsg=errmsg+"<SCRIPT language=JavaScript>alert(' ��������! ����ϵ����Ա. ');javascript:top.window.close();</SCRIPT>"
	founderr=true
end if
set rs=nothing
conn.close
set conn=nothing
if founderr=true then
	call error()
end if
%>