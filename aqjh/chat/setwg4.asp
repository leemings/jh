<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
pai=aqjh_name
rs.open "SELECT * FROM �û� where ����='"&aqjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
else
 hydj=rs("��Ա�ȼ�")
end if
%>
<html>
<head>
<title><%=pai%>��ԯ��ʽ</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (liji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" background="<%=chatimage%>" oncontextmenu=window.event.returnValue=false>
<div align="center"><font color=ffffff>��<%=pai%>�书��</font><br><br>
  <table cellpadding="3" cellspacing="0" border="1" bgcolor="#CCCCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr valign="middle"> 
      <td width="43%"> <div align="center"><font color="#330033">��ʽ</font></div></td>
      <td width="29%"> <div align="center"><font color="#330033">�书</font></div></td>
      <td width="28%"> <div align="center"><font color="#330033">����</font></div></td>
    </tr>
    <%
rs.close
rs.open "SELECT * FROM n where b='" & pai & "' order by e",conn
do while not rs.eof and not rs.bof
 if hydj<1 then
	conn.Execute "delete * from n where b='"&aqjh_name&"'"
	Response.Write "<script Language=Javascript>alert('��ʾ:�㲻�ǵȼ���Ա,����������ԯ!');parent.history.go(-1);</script>"
	response.end
 end if
dj=int(sqr((rs("e")/100)))+1
%>
    <tr valign="middle"> 
      <td width="43%"> <div align="center"> <font color="#330033"> <a href="javascript:s('/��ԯ$ <%=rs("a")%>')" title="������Ϊ��<%=dj%>��!"><%=rs("a")%></a></font></div></td>
      <td width="29%"> <div align="center"><font color="#330033"><%=rs("c")%></font></div></td>
      <td width="28%"> <div align="center"><font color="#330033"><%=rs("d")%></font></div></td>
    </tr>
	<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing%>
        </table>
</td>
    </tr>
    
  </table>
  <table cellpadding="3" cellspacing="0" border="1" bgcolor="#CCCCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr> 
            <td><div align="center"><%if hydj>0 then%><a href=# onclick="javascript:chatwin=window.open('xx/SETWG.ASP','win_aqjh','left=0,scrollbars=yes,resizable=0,width=700,height=400')">���|�޸�|����</a><%else%>�ȼ���Ա����ʹ��<%end if%></div></td>
          </tr>
        </table>
</div>
<p class=p1 align="center"><font color="#FFFFFF">ɱ��������<br>
��������+������<br>
<input type="checkbox" name="liji" value="on">
�������� </font></p>
</body></html>