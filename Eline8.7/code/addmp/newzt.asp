<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if instr(trim(request.form("subject")),"�ٸ�")<>0 or instr(trim(request.form("subject")),"������")<>0 or instr(trim(request.form("subject")),"����")<>0 then
%>
<script language=vbscript>  
MsgBox "�������ƺ��зǷ��ַ�����������д�������ƣ�"  
location.href = "javascript:history.back()"  
</script> 
<%
end if
if sjjh_name="" then %>
<script language=vbscript>  
MsgBox "�Բ������ȵ�¼������@_@"  
location.href = "javascript:history.back()"  
</script> 
<% 
end if
if trim(request.form("subject"))="" or trim(request.form("comment"))=""  then
%>
<script language=vbscript>
MsgBox "���󣺰������ƺͼ�������д��"
location.href = "javascript:history.back()"
</script>
<%
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from p Where a='"&request.form("subject")&"'",conn
if not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('��ʾ���������Ѿ����ڣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs.eof or rs.bof then 
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="select * from �û� where ����='"&Session("sjjh_name")&"' "
set rs=conn.execute(sql)
%>
<%    end if
if rs("�ȼ�")<30 or rs("����")<200 or rs("����")<200 or rs("����")<200 or rs("����")<200 or rs("����")<5000 or rs("����")<50000000 then
%>
<script language=vbscript>  
MsgBox "������û�е�¼������@_@"  
location.href = "javascript:history.back()"  
</script> 
<script language=vbscript>  
MsgBox "������Ĳ���̫�ͣ���Ҫ���ܣ�@_@"  
location.href = "javascript:history.back()"  
</script> 

<%  elseif rs("����")<>"δ��" and rs("����")<>"��" and rs("����")<>"����" then
%>
       <script language=vbscript>  
  MsgBox "�������Ѿ��������ˣ��������˳���@_@"  
  location.href = "javascript:history.back()"  
</script> 

<%  else
		str="Insert Into p (a,b,c,d,e,g,h) Values('"&request.form("subject")&"','"&sjjh_name&"',0,'"&request.form("comment")&"','��',0,0)"
		conn.execute(str)
        conn.execute "update �û� set ����='"&request.form("subject")&"',���='����',grade=5,����=����-100000000 where ����='"&sjjh_name &"'"

   end if
%>

<html>

<head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt; color:'ffff77';
	scrollbar-face-color:'222222'; 
	scrollbar-shadow-color:'222222'; 
	scrollbar-highlight-color:'EFBA56'; 
	scrollbar-3dlight-color:'222222'; 
	scrollbar-darkshadow-color: 'AB7510'; 
	scrollbar-track-color:'FBEED7'; 
	scrollbar-arrow-color:'AB7510' ;

	}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'ffff77'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff4d4d;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: '291C03'; border: 1 solid '291C03'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; 
	BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-RIGHT-WIDTH: 1px; 
	BORDER-RIGHT-COLOR:'EBA82C'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: '333333'; 
	border: 333333px solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'F3CB81'; color: '83590C'; 
	border: 1px solid AB7510;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'EBA82C'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'EBA82C'}
</style>
<style>.skin0 {
	BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; VISIBILITY: hidden; BORDER-LEFT: 1px solid; WIDTH: 80px; CURSOR: default; LINE-HEIGHT: 20px; BORDER-BOTTOM: 1px solid; FONT-FAMILY: Verdana; POSITION: absolute; BACKGROUND-COLOR: white
}
.menuitems {
	PADDING-RIGHT: 10px; PADDING-LEFT: 15px
}
                </style>



<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�Դ����ɡ�һ�������wWw.51eline.com��</title>
</head>

<body bgcolor="#000000">

<div align="center">
  <center>
  <table border="0" width="80%" cellspacing="0" cellpadding="0" height="135" background="../picc/bg.jpg">
    <tr>
      <td width="100%" colspan="3" background="../picc/k2.gif" height="18">��</td>
    </tr>
    <tr>
      <td width="4%" rowspan="4" background="../picc/k3.gif" height="99">��</td>
      <td width="88%" align="center" height="26"><img border="0" src="22.gif" width="111" height="22"></td>
      <td width="4%" background="../picc/k3.gif" height="99" rowspan="4">��</td>
    </tr>
    <tr>
      <td width="88%" height="70" align="center"><img border="0" src="../picc/king3.jpg" width="488" height="229"></td>
    </tr>
    <tr>
      <td width="88%" height="82" valign="top">
        <table border="0" width="100%">

          <tr>
            <td width="100%" align="center">��ϲ�㴴��ɹ��Ӵ����ֶ���</td>
          </tr>

          <tr>
            <td width="100%">        

              <table border="0" width="100%">
                <tr>
                  <td width="100%" align="center">�����ѳɹ���ӣ�</td>
                </tr>
              </table>
              </FORM>
            </td>
          </tr>
         <td><hr>

          <tr>
            <td width="100%" align="center">��E�߽�����</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="88%" height="1"></td>
    </tr>
    <tr>
      <td width="100%" colspan="3" background="../picc/k2.gif" height="18">��</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>






















































