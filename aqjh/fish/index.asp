<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
myid= Replace(Session("aqjh_name"), "'", "''")
if myid="" then
	Response.Write("�������ʺš�")
	Response.End
End If
err_num=0
err_info=""

Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Open Application("aqjh_usermdb")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs1 = Server.CreateObject("ADODB.Recordset")

Sql_1="SELECT * FROM �û� WHERE ���� = '" & myid & "'"
Rs.Open Sql_1,Conn,1,3
If Rs.Bof and Rs.Eof Then 'û�ҵ� ��Ʒ�ǿ�
	err_num=err_num+1
	err_info=err_info+"�ʺŲ�����"
	jh_fish=0
Else '�ҵ� ��Ʒ����
	Rs_numRows = 0
	jh_fish=Rs.Fields.Item("jh_fish").Value
	if jh_fish="" or isnull(jh_fish) then jh_fish=0
	dj="����"
	if jh_fish>=300 then dj="���㹤��"
	if jh_fish>=600 then dj="���㼼ʦ"
	if jh_fish>=2000 then dj="����ר��"
	if jh_fish>=20000 then dj="�����ʦ"
	
	
	fish_date=Rs.Fields.Item("����ʱ��").Value 
	'Response.Write(DateDiff("s",fish_date,now()))
	fish_sj=DateDiff("s",fish_date,now())
	if fish_sj<120 then
	Set Rs = Nothing
	Set Rs1 = Nothing
	Conn.Close()
	Set Conn = Nothing
	Response.Write("<script language=""JavaScript"">alert('��ʾ����������ÿ��20��,������������...\n\nʱ�仹û��,����"&(120-fish_sj)&"��')</script>")
	Response.end
	end if
End If 
Rs.Close()
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������</title>
<link href="fish.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/JavaScript">
<!--
var tool1=0
function load_div(text,obj){
   var x = 0;   var y = 0;  var i = 0;
   tab_01.rows[0].cells[0].innerHTML= text;
   while (obj!=null) {x=0;y=0;if (i>2) {x+=obj.offsetLeft; y+=obj.offsetTop;}i++; obj = obj.offsetParent; }
   div_del.style.left=x+480;//   div_del.style.top=y;
   div_del.style.visibility='visible';
   return false;
}
function formSubmit(){
if (true==edit_c()) document.form.submit();}
function edit_c(){
 if (document.form.address.value==''){alert("ûѡ�ص�"+document.form.address.value);return false;}
 if (document.form.tool.value==''){alert("ûѡ���"+document.form.tool.value);return false;}
 if (document.form.bait.value==''){alert("ûѡ���"+document.form.bait.value);return false;}
 if (tool_n = 0){alert("�ص����߲�����");return false;}
 if (<%= err_num %>==0) {}else{alert("���ݳ���");return false;}
 return true;}
function edit_a(address,tool,tool_num){
document.form.address.value=address;
document.form.tool.value=tool;
tool_n=tool_num
}
function edit_b(bait){
document.form.bait.value=bait;
}
//-->
</script>
<link href="base.css" rel="stylesheet" type="text/css" />
</head>
<body> 
<div id="div_del" style="position:absolute; left:450px; top:68px; width:65px; height:57px; z-index:1; visibility: hidden;"> 
  <table width="280" border="0" cellpadding="4" cellspacing="0" class="table_bk" id="tab_01"> 
    <tr> 
      <td height="53" colspan="3"><b>Զ��</b><br> 
        �������ˣ����ض��顣<br> 
        ������ߣ��������(����-������-������)<br> 
        ��ѡ�������Σ�С��</td> 
    </tr> 
  </table> 
</div> 
<table width="760" border=0 align="center" cellpadding="0" cellspacing="0" bgcolor="#E8E8F0"> 
  <tr> 
    <td width="167" valign="top"><br> 
      <table width="167" border="0" cellspacing="0" cellpadding="0" height="80"> 
        <tr> 
          <td>&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/cloud02.gif" width="60" height="24"></td> 
          <td><br> 
            <br> 
            <img src="images/cloud03.gif" width="34" height="14"></td> 
        </tr> 
      </table></td> 
    <td> <img src="images/ship01.gif" width="129" height="46"><img src="images/ship02.gif" width="137" height="95" border="0"></td> 
    <td valign=top><img src="images/ship03.gif" width="50" height="60"></td> 
    <td width="100%"> <table width="100%" border="0" cellspacing="0" cellpadding="0" height="90"> 
        <tr> 
          <td>&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/cloud02.gif" width="60" height="24"></td> 
          <td valign="bottom"><img src="images/cloud03.gif" width="33" height="14"><br> 
            <br></td> 
          <td><img src="images/cloud01.gif" width="69" height="28"><br /> 
            <br></td> 
        </tr> 
      </table></td> 
  </tr> 
</table> 
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#d3d7ea"> 
  <tr bgcolor="#333333"> 
    <td colspan="4" valign="top"><img src="empty.gif" width="1" height="2"></td> 
  </tr> 
  <tr> 
    <td width="174" valign="top"><img src="empty.gif" width="174" height="1"></td> 
    <td width="30"><img src="images/ship04.gif" width="30" height="38"></td> 
    <td width="161" valign="top"><img src="images/ship05.gif" width="161" height="15"></td> 
    <td bordercolor="#FFFFFF">��</td> 
  </tr> 
</table> 
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="bcd3d7ea" id="tab_main"> 
  <tr> 
    <td><div align="center"><br />
        <br /><%= Trim(Request.QueryString("info")) %>
        <br /> 
        <br />
    </div></td> 
  </tr> 
  <tr> 
    <td> <table width="660" border="0" align="center" cellpadding="0" cellspacing="0"> 
        <form name="form" id="form" method="post" action="fish1/index.asp"> 
          <tr> 
            <td><table width="540" border="0" cellspacing="0" cellpadding="0" align="center"> 
                <tr> 
                  <td class="f14 fb">�ڵ���ǰ������ѡ��ô����ص�����������
                    <input name="address" type="hidden" value="" /> 
                    <input name="bait" type="hidden" /> 
                    <input name="tool" type="hidden" /> 
                    <br /> 
                    <br /></td> 
                </tr> 
                <tr> 
                  <td class="ct-def1">����ǰ�ļ�����<font color="#cc3366"><%= dj %><%= err_info %></font>��(������Ƶ������ϣ����Կ�����صĽ���)</td> 
                </tr> 
              </table></td> 
          </tr> 
          <tr> 
            <td> </td> 
          </tr> 
          <tr> 
            <td> <table border="0" align="center" cellpadding="4"> 
                <tr> 
                  <th>�ص�ѡ��</th> 
                  <th>�ر����</th> 
                  <th>��ѡ���</th> 
                  <th>���ѡ��</th> 
                </tr> 
<% 
Sql_1="SELECT * from aqjh_fish"
Rs.Open Sql_1,Conn,0,1

While (NOT Rs.EOF) 
' jh_fish
dy_dj=Rs.Fields.Item("dy_dj").Value
fish_address=Rs.Fields.Item("fish_address").Value
fish_tool=Rs.Fields.Item("fish_tool").Value
fish_bait=Rs.Fields.Item("fish_bait").Value

Rs1.open "select * from wpname where wp_user='" & myid & "' and wp_name='"& Replace(fish_tool, "'", "''") &"'",conn
If Rs1.Bof and Rs1.Eof Then 
tool_num=0
Else
tool_num=Rs1.Fields.Item("wp_sl").Value
End If
Rs1.Close
Rs1.open "select * from wpname where wp_user='" & myid & "' and wp_name='"& Replace(fish_bait, "'", "''") &"'",conn
If Rs1.Bof and Rs1.Eof Then 
bait_num=0
Else
bait_num=Rs1.Fields.Item("wp_sl").Value
End If
Rs1.Close %> 
                <tr> 
                  <td class="co<% If jh_fish>=dy_dj Then Response.Write("1")%>"><% If jh_fish>=dy_dj Then %><input name="daddress" onClick="edit_a(<% Response.Write("'"&fish_address&"','"&fish_tool&"',"&tool_num) %>)" type="radio" /><% End If %>
				  <a href="#"  onmouseover="javascript:load_div('<%= Rs.Fields.Item("fish_main").Value %>',this)" onMouseOut="javascript:div_del.style.visibility='hidden'"><%= fish_address %></a></td> 
                  <td class="co<% If tool_num>1 Then Response.Write("1")%>"><% Response.Write(Rs.Fields.Item("fish_tool").Value)&tool_num %></td> 
                  <td><%=(Rs.Fields.Item("fish_all_bait").Value)%></td> 
                  <td class="co<% If bait_num>0 Then Response.Write("1")%>"><% If bait_num>0 Then %><input name="dbait" onClick="edit_b(<%= "'"&fish_bait&"'" %>)" type="radio" /> <% End If %><% Response.Write(Rs.Fields.Item("fish_bait").Value&bait_num) %></td> 
                </tr> 
<% Rs.MoveNext()
Wend 
Rs.Close()
Set Rs = Nothing
Set Rs1 = Nothing
Conn.Close()
Set Conn = Nothing
%> 
              </table></td> 
          </tr> 
          <tr> 
            <td><table width="540" border="0" cellspacing="0" cellpadding="0" align="center"> 
                <tr> 
                  <td colspan="2"><span class="ct-def1"><b>ע��Ŷ������</b></span><br /> 
                    <br /> 
                    <span class="ct-def1">�ص���<font color="#93abe3">ǳ��</font>��ǣ�˵����Ŀǰ�ļ�����޷�ȥ�õص���㣬ҪŬ��Ŷ��<br /> 
                    ��߻������<font color="#93abe3">ǳ��</font>��ǣ�˵�����ı������ﻹû����������Ʒ���ɵ����ߵ�ȥ����<br /> 
                    <br /> 
                    </span></td> 
                </tr> 
                <tr> 
                  <td><img src="images/ok.gif" border="0" width="36" height="36" align="absmiddle" /> <a href="wplist.htm" class="f14">ȥ����Ʒ�̵깺�������Ʒ</a></td> 
                  <td align="right"><img src="images/ok.gif" border="0" width="36" height="36" align="absmiddle" /> <a href="javascript:formSubmit()" class="f14">������</a> 
                    <input type="hidden" name="ok" value="0" /></td> 
                </tr> 
              </table></td> 
          </tr> 
        </form> 
      </table></td> 
  </tr> 
  <tr> 
    <td><br /> </td> 
  </tr> 
</table> 
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#d3d7ea"> 
  <tr> 
    <td width="59"><img src="images/botton_01.gif" width="59" height="121" /></td> 
    <td width="700" valign="bottom"><img src="images/botton_02.gif" width="700" height="73" /></td> 
  </tr> 
</table> 
</body>
</html>


</body>
</html>