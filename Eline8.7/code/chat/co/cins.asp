<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
rs.Open ("select count(*) from �û� where ���='��ʬ��'"),conn
count=rs(0)
rs.Close
rs.Open ("select top 1 ���� from �û� where ���='��ʬ��'"),conn
sql="Select  * from �û� where ����='"&sjjh_name&"'"  
set rs=conn.Execute(sql)
attackname=rs("����")
sf=rs("���")
%>

<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��ʬ�ͷš�wWw.happyjh.com��</title>
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt; color:'ffff00';
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
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
	PADDING-CENTER: 1px; PADDING-CENTER: 1px; 
	BORDER-CENTER-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-CENTER-WIDTH: 1px; 
	BORDER-CENTER-COLOR:'EBA82C'}

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

</head>

<body bgcolor="#006699" topmargin="2" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>

  <table width="140" height="53" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#339999">
 
    <tr>
      <td width="100%" height="47" valign="top">
        <div align="center">
          <center>
          <table border="0" width="100%" cellpadding="0" height="56">
            <tr>
              <td width="103%" colspan="2" align="center" height="20">����ʬ���</td>
            </tr> 
            <tr>
              <td width="108%" align="center" height="12" colspan="2"><hr color="#FFFFCC" size="1"></td>
            </tr> 
            <tr>
              <td width="101%" align="center" height="12" colspan="2">����<%=count%>ֻ��ʬ��������</td>
               </tr>  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� where ���='��ʬ��' order by ���� ",conn
attackname=rs("����")
sf=rs("���")
%>
           <tr>
              <td width="33%" align="center" height="12">��ʬ��</td>
              <td width="34%" align="center" height="12">����</td>
               </tr>
               
      <%
do while not rs.bof and not rs.eof
%>
 <tr>
              <td width="33%" align="center" height="10">
              <a  href="#" onClick="window.open('../../yamen/mt.asp?action=<%=rs("����")%>','wupin','scrollbars=yes,resizable=yes,width=550,height=500')" title="�鿴״̬��"><%=rs("����")%></font></a>
</td>
     <td id="sm" width="34%" align="center" height="10"><a href=../attack.asp?id=<%=rs("����")%>&value=100000 target="d" title="�ٸ������ʸ��ͷ�">�ͳ�</td>
                       </tr> 
                                                                                   
    <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

   
            <tr>

              <td width="103%" height="14" colspan="2" align="center"><hr color="#FFFFCC" size="1"></td>
            </tr>


            <tr>
              <td width="103%" height="14" colspan="2" align="center">��ʬ������ǿ��</td>
            </tr>

            <tr>
              <td width="103%" height="14" colspan="2" align="center">�ٸ������ʸ��ͷ�</td>
            </tr>
 

            <tr>
              <td width="103%" height="14" colspan="2" align="center"><%=Application("sjjh_chatroomname")%>�״�</td>
            </tr>

          </table>
            </center>
        </div>
     </td>
    </tr>
  </table>

</body>

</html>

















































































































































































