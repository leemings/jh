<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
myid= Replace(Session("aqjh_name"), "'", "''")
if myid="" then
	Response.Write("请输入帐号。")
	Response.End
End If
err_num=0
err_info=""

Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Open Application("aqjh_usermdb")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs1 = Server.CreateObject("ADODB.Recordset")

Sql_1="SELECT * FROM 用户 WHERE 姓名 = '" & myid & "'"
Rs.Open Sql_1,Conn,1,3
If Rs.Bof and Rs.Eof Then '没找到 物品是空
	err_num=err_num+1
	err_info=err_info+"帐号不存在"
	jh_fish=0
Else '找到 物品存在
	Rs_numRows = 0
	jh_fish=Rs.Fields.Item("jh_fish").Value
	if jh_fish="" or isnull(jh_fish) then jh_fish=0
	dj="新手"
	if jh_fish>=300 then dj="捕鱼工人"
	if jh_fish>=600 then dj="捕鱼技师"
	if jh_fish>=2000 then dj="捕鱼专家"
	if jh_fish>=20000 then dj="捕鱼大师"
	
	
	fish_date=Rs.Fields.Item("操作时间").Value 
	'Response.Write(DateDiff("s",fish_date,now()))
	fish_sj=DateDiff("s",fish_date,now())
	if fish_sj<120 then
	Set Rs = Nothing
	Set Rs1 = Nothing
	Conn.Close()
	Set Conn = Nothing
	Response.Write("<script language=""JavaScript"">alert('提示：钓鱼限制每次20秒,请您过会再来...\n\n时间还没到,还有"&(120-fish_sj)&"秒')</script>")
	Response.end
	end if
End If 
Rs.Close()
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>江湖钓鱼</title>
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
 if (document.form.address.value==''){alert("没选地点"+document.form.address.value);return false;}
 if (document.form.tool.value==''){alert("没选鱼竿"+document.form.tool.value);return false;}
 if (document.form.bait.value==''){alert("没选鱼饵"+document.form.bait.value);return false;}
 if (tool_n = 0){alert("地点和鱼具不符合");return false;}
 if (<%= err_num %>==0) {}else{alert("数据出错");return false;}
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
      <td height="53" colspan="3"><b>远洋</b><br> 
        搏击风浪，满载而归。<br> 
        所需渔具：超级渔具(缆绳-捕鲸炮-捕鲸叉)<br> 
        可选鱼饵：猪肝，小鱼</td> 
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
    <td bordercolor="#FFFFFF">　</td> 
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
                  <td class="f14 fb">在钓鱼前，请您选择好垂钓地点和所需鱼饵：
                    <input name="address" type="hidden" value="" /> 
                    <input name="bait" type="hidden" /> 
                    <input name="tool" type="hidden" /> 
                    <br /> 
                    <br /></td> 
                </tr> 
                <tr> 
                  <td class="ct-def1">您当前的级别是<font color="#cc3366"><%= dj %><%= err_info %></font>　(将鼠标移到地名上，可以看到相关的介绍)</td> 
                </tr> 
              </table></td> 
          </tr> 
          <tr> 
            <td> </td> 
          </tr> 
          <tr> 
            <td> <table border="0" align="center" cellpadding="4"> 
                <tr> 
                  <th>地点选择</th> 
                  <th>必备渔具</th> 
                  <th>可选鱼饵</th> 
                  <th>鱼饵选择</th> 
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
                  <td colspan="2"><span class="ct-def1"><b>注意哦！！！</b></span><br /> 
                    <br /> 
                    <span class="ct-def1">地点以<font color="#93abe3">浅蓝</font>标记，说明您目前的级别可无法去该地点钓鱼，要努力哦。<br /> 
                    渔具或鱼饵以<font color="#93abe3">浅蓝</font>标记，说明您的宝物箱里还没有这样的物品，可到工具店去购买。<br /> 
                    <br /> 
                    </span></td> 
                </tr> 
                <tr> 
                  <td><img src="images/ok.gif" border="0" width="36" height="36" align="absmiddle" /> <a href="wplist.htm" class="f14">去工具品商店购买钓鱼用品</a></td> 
                  <td align="right"><img src="images/ok.gif" border="0" width="36" height="36" align="absmiddle" /> <a href="javascript:formSubmit()" class="f14">出发啦</a> 
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