<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���,����,����,�书,����,����,�ȼ�,����,����,���� from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=440"
	response.end
end if
if rs("�ȼ�")<200 or rs("����")<500000 or rs("����")<1500000 or rs("�书")<550000 or rs("����")<340600 or rs("����")<336700 or rs("����")<50000 or rs("����")<10000 or rs("����")<500000000 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����״̬δ�ﵽ�������뿴����������');window.close();</script>"
	response.end
end if
if rs("���")="����" then
	Response.Write "<script Language=Javascript>alert('��ʾ���������Ѿ��������ˣ�����Ҫ��������ѽ��');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("����")="�ٸ�" or (rs("����")<>"��" and rs("����")<>"����") then
	Response.Write "<script Language=Javascript>alert('��ʾ��Ҫ���Դ����ɾͱ��������뿪���ڵ����ɣ�');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
tmprs=conn.execute("Select count(*) As ���� from �û� where �ȼ�>=80 and ������='"& aqjh_name &"'")
lr=tmprs("����")
set tmprs=nothing
if lr<8 then
	Response.Write "<script Language=Javascript>alert('��ʾ��������˼�¼����8�ˣ������������˵ĵȼ���û�е�80�����ϣ�');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
%>
<html>
<head>
<title>����</title>
<link rel="stylesheet" href="../../css.css">
</head>
<BODY oncontextmenu=self.event.returnValue=false bgColor=#339966 text=#ffffff>
<p align="center"><img border="0" src="22.GIF" width="111" height="22"></p>
<table width="363" border="1" cellspacing="0" cellpadding="3" align="center" height="15" bordercolorlight="#999999" bordercolordark="#FFFFFF">
  <tr align="center"> 
    <td height="35">��ϲ�����������Ż�</td>
  </tr>
  <tr> 
    <td height="108"> 
      <form method="post" action="newzt.asp">
        <p align="center">
          <table width="77%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
            <tr> 
              <td height="36" align="center">�š����ɣ�</td>
              <td height="36"> 
                <input type="text" name="subject" size="10" maxlength="4">
              </td>
            </tr>
            <tr>
              <td align="center">���ɼ�飺</td>
              <td>
                <input type="text" name="comment" size="20" maxlength="20">
              </td>
            </tr>
          </table>
        <p align="center">
          ���ɼ����಻�ܳ���50�֡�
        </p>
        <p align="center">
          <input type="submit" name="Submit" value="ȷ��">
          <input type="reset" name="reset" value="ȡ��"> 
        </p>
      </form>
 </td>
  </tr>  
</table>
</body>
</html>