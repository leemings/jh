<!--#include file="conn.asp"-->
<html>
<title>���ֽ�����վ___DJ���___wWw.happyjh.com</title>
<body style='background:transparent'>
<script src="js/js.js"></script>
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
	COLOR: #333333
}
A:active {
	COLOR: #FF0000
}
A:hover {
	COLOR: #333333; TEXT-DECORATION: underline overline
}


</style>

<%
if (isnull(session("MusicUser")) and isnull(session("MusicPwd")) and session("MusicUser")="" and session("MusicPwd")="") and Session("IsAdmin")=true and session("KEY")="super" then
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('Ϊ�����ע���û��������Ҫ��ɧ�ţ����Ի�Ա��Ϣ���Ա�վע���û�����!\n������Ѿ�ע�ᣬ���ȵ�½Ȼ���ٲ鿴�������û��ע�ᣬ����ע��!');window.close();</script>"
	Response.End 
end if
if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
%>
<div align="center">
  <center>
  <table border="1" cellspacing="0" bordercolor="#666666" width="100%" id="AutoNumber1" height="1" style="border-collapse: collapse" cellpadding="0" bgcolor="#FFCC66">
    <tr>
      <td width="755" height="1" bgcolor="#EC9F00">
      <p align="center">-==��վ��Ա�б�==-</td>
    </tr>
    <tr>
      <td width="100%" height="1">
      <div align="center">
        <center>
     <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td width="70%" valign="top" align="center">


��</td>
          </tr>
        </form>



</td>
  </tr>
  <tr>
    <td width="70%" valign="top" align="center">
<%
set rs=server.createobject("adodb.recordset")
sql="select * from user where username<>'���ž���' order by id desc" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
	response.write "<p align='center'>��ʱû���û�ע��</p>" 
else 
	MaxPerPage=18
	PageUrl="UserList.asp"
	totalPut=rs.recordcount 
	if currentpage<1 then currentpage=1
	if (currentpage-1)*MaxPerPage>totalput then 
		if (totalPut mod MaxPerPage)=0 then 
			currentpage= totalPut \ MaxPerPage 
		else 
			currentpage= totalPut \ MaxPerPage + 1 
		end if 
	end if 
	if currentPage=1 then 
		showpage totalput,MaxPerPage,PageUrl
		showContent 
		showpage totalput,MaxPerPage,PageUrl
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move  (currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark 
			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		else 
			currentPage=1 
			showpage totalput,MaxPerPage,PageUrl
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		end if 
	end if 
end if 
rs.close 
			
sub showContent
i=0
%>
            <table border="1" width="100%" cellspacing="0" bordercolor="#666666" height="1" cellpadding="0" style="border-collapse: collapse">
              <tr>
                <td width="28" height=1 align=center bgcolor="#EC9F00">
                ID</td>
                <td width="94" height=1 align=center bgcolor="#EC9F00">
                �û���</td>
                <td width="36" height=1 align=center bgcolor="#EC9F00">
                �Ա�</td>
                <td width="155" height=1 align=center bgcolor="#EC9F00">
                ��������</td>
                <td width="72" height=1 align=center bgcolor="#EC9F00">
                OICQ</td>
                <td width="68" height=1 align=center bgcolor="#EC9F00">
                ��ҳ</td>
                <td width="136" height=1 align=center bgcolor="#EC9F00">
                ������Ϣ</td>
              </tr>
<%
do while not rs.eof
	i=i+1
	id=rs("id")
%>
              <tr>
                <td align=center width="28" height="1"><%=id%></td>
                <td align=center width="94" height="1"><a href="javascript:open_window('UserInfo.asp?id=<%=rs("id")%>','UserInfo','scrollbars=no,resizable=no,toolbar=no,location=no,directories=no,status=no,menubar=no,copyhistory=no,top=10,left=10,width=300,height=200')"><%=rs("UserName")%></a><%if DateValue(rs("regDate"))=date() then%> <img src="images/new.gif"><%end if%>
<%if Session("IsAdmin")=true and session("KEY")="super" then%>
<br>
<a href="cnadmin/UserModify.asp?id=<%=rs("id")%>"><font color="#FF0000">��</font></a><font color="#FF0000">
                </font>
<a href="cnadmin/UserDel.asp?id=<%=rs("id")%>"><font color="#FF0000">ɾ</font></a><font color="#FF0000">
                </font>
<%end if%>
                </td>
                <td align=center width="36" height="1"><%if rs("sex")=true then%>��<%else%>��<%end if%></td>
                <td align=center width="155" height="1"><%=rs("Email")%></td>
                <td align=center width="72" height="1"><%if rs("oicq")<>"" then%><%=rs("oicq")%><%else%>��<%end if%></td>
                <td align=center width="68" height="1"><%if rs("homepage")<>"" then%><a target="_blank" href="http://<%=rs("homepage")%>">��˲鿴</a><%else%>��<%end if%></td>
                <td align=center width="136" height="1">
                <a href="javascript:open_window('messager.asp?action=new&touser=<%=rs("UserName")%>','messanger','width=420,height=290')"><font color="#FF0000">
          ��</font> 
            <%=rs("UserName")%> <font color="#FF0000">������Ϣ</font></a></td>
              </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
            </table>
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<form method=Post action="<%=filename%>">
  <center>
  <p>��<font color="red"><b><%=totalnumber%></b></font>λ�û�
<%if CurrentPage<2 then%> &nbsp;��ҳ ��һҳ&nbsp;
<%else%> &nbsp;<a href="<%=filename%>?page=1">��ҳ</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>">
  ��һҳ</a>&nbsp;
<%
end if
if n-currentpage<1 then
%> ��һҳ ĩҳ
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>">
  ��һҳ</a>
  <a href="<%=filename%>?page=<%=n%>">
  ĩҳ</a>
<%end if%> &nbsp;ҳ��:<strong><font color="red"><%=CurrentPage%>/<%=n%></font></strong>ҳ 
  ת��:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>
  ��<%=i%>ҳ</option>   
<%next%>   
  </select> </p>
</form>        
<%end function%>
    </td>
  </tr>
  </table></center>
      </div>
      </td>
    </tr>
    <tr>
      <td width="750" height="7" bgcolor="#EC9F00">
      &nbsp;
      </td>
    </tr>
  </table>
  </center>
</div>
<%
set rs=nothing
conn.close
set conn=nothing%></body></html>