<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("select count(*) from �û� where ���='��ʬ��' and ״̬='��ʬ��'"),conn
count=rs(0)
rs.Close
rs.Open ("select top 1 ���� from �û� where ���='��ʬ��'"),conn
sql="Select  * from �û� where ����='"&aqjh_name&"'"  
set rs=conn.Execute(sql)
attackname=rs("����")
sf=rs("���")
%>
<html><head>
<meta http-equiv="Content-Language" content="zh-tw">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��ʬ�ͷ�</title>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../bg.gif" bgproperties="fixed">
  <table border="1" width="100%" cellspacing="0" cellpadding="2" bordercolordark="#FFFFFF" bordercolorlight="#000000" background="../BG.gif" height="53">
 
    <tr>
      <td width="100%" background="../BG.gif" height="47" valign="top">
        <div align="center">
          <center>
          <table border="0" width="100%" cellpadding="0" height="56">
            <tr>
              <td width="100%" colspan="2" align="center" height="20">����ʬ���</td>
            </tr> 
            <tr>
              <td width="100%" align="center" height="12" colspan="2"><hr color="#FFFFCC"></td>
            </tr> 
            <tr>
              <td width="100%" align="center" height="12" colspan="2">����<%=count%>ֻ��ʬ��������</td>
               </tr>  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if count>0 then
rs.open "SELECT * FROM �û� where ״̬='��ʬ��' and ���='��ʬ��' order by ���� ",conn
attackname=rs("����")
sf=rs("���")
else
response.end
end if
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
     <td id="sm" width="34%" align="center" height="10"><a href=attack.asp?id=<%=rs("����")%>&value=100000 target="d" title="����2�������ʸ��ͷ�">�ͳ�</td>
                       </tr> 
                                                                                   
    <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%><tr> 
<td width="103%" height="14" colspan="2" align="center"><hr color="#FFFFCC"></td>
</tr><tr><td width="103%" height="14" colspan="2" align="center">��ʬ������ǿ��</td>
</tr><tr><td width="103%" height="14" colspan="2" align="center">����ȼ�2�����ʸ��ͷ�</td>
</tr></table></center></div></td></tr></table></body></html>