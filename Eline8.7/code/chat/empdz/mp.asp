<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="SELECT ����,���� FROM �û� where ����='" & sjjh_name &"'"
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "�㲻�ǽ�������"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
else
	pai=rs("����")
	tl=rs("����")
rs.close
sql="SELECT q,s FROM p where a='"&pai&"'"
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	shengming="û������ս����"
	shidi="û�е���"
	else
	shengming=rs("s")
	shidi=rs("q")
	end if
rs.close
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���ɴ�ս</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (liji==true) {parent.f2.document.af.subsay.click()};
}
</script>
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
	
INPUT {
	BACKGROUND: #F5E1B9; BORDER-BOTTOM: #4E390F 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #4E390F 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #4E390F; CURSOR: hand; FONT-FAMILY: "����"; FONT-SIZE: 12px; HEIGHT: 20px; PADDING-BOTTOM: 1px; PADDING-LEFT: 1px; PADDING-RIGHT: 1px; PADDING-TOP: 1px

}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'003300'; font-size: 9pt;}
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

<body bgcolor="#006699" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<div align="right">
  <table border="1" width="100%" cellspacing="0" cellpadding="2" bgcolor="#33CCCC" bordercolordark="#FFFFFF" bordercolorlight="#000000" align="right">
    <tr>
      <td width="100%" align="center" bgcolor="#66CCCC"><img border="0" src="ki17.gif" width="18" height="18"><font color="#CC0000">��<strong>���ɴ�ս</strong>��</font><img border="0" src="ki17.gif"></td>
    </tr>
    <tr>
      <td width="100%"><a href="mpbao.asp">���뱣��</a>:<font color="#FFFFCC">�������ɱ���</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="mpbao2.asp">�رձ���</a>:<font color="#FFFFCC">�ر����ɱ���</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="javascript:s('/������ս$')">������ս</a>:<font color="#FFFFCC">����������ս</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="javascript:s('/�������$')">�������</a>:<font color="#FFFFCC">������������</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="mp2.asp">����ս��</a>:<font color="#FFFFCC">��������ս��</font></td>
    </tr>
    <tr>
      <td width="100%"><font color="#003300">����ս��</font>:<font color="#FFFFCC" size="2"><%=tl%></font></td>
    </tr>
    <tr>
      <td width="100%"><font color="#003300">����ս��</font>:<font color="#FFFFCC" size="2"><%=shengming%></font></td>
    </tr>
    <tr>
      <td width="100%"><font color="#003300">��������</font>:<input type="text" size="12" value="<%=shidi%>" readonly></td>
    </tr>
    <tr>
      <td width="100%" align="center"> <table cellpadding="3" cellspacing="0" border="0" align="center" width="100%">
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"><font color="#330033" size="2">��ʽ</font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font color="#330033" size="2">�书</font></font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font color="#330033" size="2">����</font></font></div>
      </td>
    </tr>
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM y where b='" & pai & "' order by e",conn
do while not rs.eof and not rs.bof
dj=int(sqr((rs("e")/100)))+1
%>
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"> <font color="#ff0000" size="2"> <a href="javascript:s('/���ɴ�ս$ <%=rs("a")%>')" title="������Ϊ��<%=dj%>��!"><%=rs("a")%></a></font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font color="#FFFFCC" size="2"><%=rs("c")%></font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font color="#FFFFCC" size="2"><%=rs("d")%></font></div>
      </td>
    </tr>
   

 
 <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
end if%>

  </table>
</td>
    </tr>
    <tr>
      <td width="100%" align="center"><div align="center"><img src="../go.gif" width="134" height="98"></div></td>
    </tr>
    <tr>
      <td width="100%" align="center"><div align="center"><font color="#FFFFFF" size="2">ɱ��������<br>
��������+������<br>
<input type="checkbox" name="liji" id="liji" value="on">
�������� </font></div></td>
    </tr>
  </table>
 
</div>
</body>

</html>

















