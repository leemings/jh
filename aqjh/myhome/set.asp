<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if session("myroom")<>"h_ӵ����" then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ǵ˷���������Ȩ���ã�');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
%>
<html><title>С�����ò���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center">
<tr>
<td align="center"> <p><img src="IMAGES/home1.gif" width="196" height="163"></p></td>
</tr>
<tr>
<td align="center">С�����ò���</td>
</tr>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT h_ӵ����2,h_�ι� FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���㻹û�з��ݻ�������Ȩ���ã�');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
peiou=rs("h_ӵ����2")
guan=rs("h_�ι�")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<tr><form name="form" method="post" action="setok.asp" >
<td align="center">
<table width="96%" border="0">
          <tr> 
            <td align="center"> �Ƿ�������ż���룺 
              <select name="peiou" id="peiou">
                <option value="1" <%if peiou<>"��" then%>selected<%end if%>>����</option>
                <option value="0" <%if peiou="��" then%>selected<%end if%>>��ֹ</option>
              </select> <br>
              �Ƿ�������˲ιۣ�              
              <select name="guan" id="guan">
                <option value="1" <%if guan=1 then%>selected<%end if%>>����</option>
                <option value="0" <%if guan=0 then%>selected<%end if%>>��ֹ</option>
              </select>
              <br>
              <br>
              <br></td>
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

