<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
'if session("myroom")<>"h_ӵ����" then
	'Response.Write "<script Language=Javascript>alert('��ʾ����Ǵ˷���������ȫ������');parent.email.style.visibility='hidden';</script>"
	'Response.End
'end if
%>
<html><title>����Ǯ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center">
<tr>
<td align="center"> <p><img src="IMAGES/money1.gif" width="244" height="191"></p></td>
</tr>
<tr>
<td align="center">ֻ�й�ͬ�ĸ���Ŭ���Ż��и�����ջ�!</td>
</tr>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT h_����Ǯ�� FROM house WHERE "& session("myroom") &"='" & aqjh_name &"'",conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���㻹û�з�����ȥ����');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
money1=rs("h_����Ǯ��")
rs.close
rs.open "SELECT ��ż,���� from [�û�] WHERE ����='" & aqjh_name & "'",conn,1,1
'mywife=rs("��ż")
myyin=rs("����")
rs.close
'rs.open "SELECT ���� from [�û�] WHERE ����='" & mywife & "'",conn,1,1
'if rs.Eof or rs.Bof then
	'rs.close
	'set rs=nothing
	'conn.close
	'set conn=nothing
	'Response.Write "<script Language=Javascript>alert('��ʾ���Ҳ��������ż����!');parent.email.style.visibility='hidden';</script>"
	'response.end
'end if
'mywifezh=rs("����")
'rs.close
'rs.open "SELECT h_����Ǯ�� FROM house WHERE h_ӵ����='" & mywifezh &"'",conn,1,1
'if rs.Eof or rs.Bof then
	'rs.Close
	'Set rs = Nothing
	'conn.Close
	'Set conn = Nothing
	'Response.Write "<script Language=Javascript>alert('��ʾ�������ż��û�з��ݣ��㻹��������һ����');parent.email.style.visibility='hidden';</script>"
	'Response.End
'end if
'money2=rs("h_����Ǯ��")
'rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<tr><form name="form" method="post" action="money1.asp" onsubmit='return(check());'>
<td align="center">
<table width="96%" border="0">
          <tr> 
            <td align="center"> ��Ҫ: 
              <select name="fs" id="fs">
                <option value="1">��Ǯ���</option>
                <option value="2">ȡǮ���</option>
              </select> <input name="money" type="text" id="money" size="10" maxlength="10" value="10000">
              ��<br>
              �����ֽ�<%=myyin%>��<br>
              �ҵ�Ǯ����: 
              <input name="money1" type="text" readonly id="money1" size="10" maxlength="10" value="<%=money1%>">
              �� </td>
          </tr>
          <tr>
            <td align="center">
<input type="submit" name="Submit" value="ȷ��">
            </td>
          </tr>
        </table>
</td>
</form>
</tr>

</table>
<SCRIPT language=JavaScript>
var sl=0;
function check(){
var pattern = /^([0-9])+$/;
var money1=parseFloat(document.form.money1.value);
var money=document.form.money.value;
var fs=parseFloat(document.form.fs.value);
if(pattern.test(money)!=true){alert("��ʾ�����������֣�");return false;}
money=parseFloat(money);
if (fs==1 && money<10000){alert("��ʾ���������е�Ǯ����������1������");return false;}
if (fs==2 && money<1000){alert("��ʾ��ȡ����Ǯ����������1000����");return false;}
if (fs==2 && money>money1){alert("��ʾ����Ľ����û����ô���Ǯ��");return false;}
return true; 
}
</SCRIPT>






