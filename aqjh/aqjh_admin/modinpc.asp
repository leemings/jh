<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="check.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
id=clng(trim(Request.QueryString("ID")))
rs.open "Select * from npc where id=" & Id,conn,1,1
conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','�����¼','�޸�NPC���[" & id & "]')"
%>
<html><head><title>�޸�NPC����</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><link href="jh.css" rel="stylesheet" type="text/css"></head>
<SCRIPT language=JavaScript>
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.form.bds.value;
revisedTitle = addTitle;
document.form.bds.value=revisedTitle;
document.form.bds.focus();
return; }
function check(){
var npcname=document.form.npcname.value;
var npcsex=document.form.npcsex.value;
var npcvalue=document.form.npcvalue.value;
var npcimg=document.form.npcimg.value;
var npcmoney=document.form.npcmoney.value;
var npcgj=document.form.npcgj.value;
var npcfy=document.form.npcfy.value;
var npcwg=document.form.npcwg.value;
var npctl=document.form.npctl.value;
var npcnl=document.form.npcnl.value;
var npcmf=document.form.npcmf.value;
var npcgjl=document.form.npcgjl.value;
var npczh=document.form.npczh.value;
var npccxl=document.form.npccxl.value;
var npcdiren=document.form.npcdiren.value;
var npcqjia=document.form.npcqjia.value;
var npcdie=document.form.npcdie.value;
var npcwp=document.form.npcwp.value;
var npczs=document.form.npczs.value;
var npcccc=document.form.npcccc.value;
var npczhuren=document.form.npczhuren.value;
var npcwplx=document.form.npcwplx.value;
if(npcname=="" ){alert("��ʾ��NPC�����ֲ���Ϊ�գ�");return false;}
if(npcsex=="" ){alert("��ʾ��NPc�Ա���Ϊ�գ�");return false;}
if(npcimg=="" ){alert("��ʾ��NPcͼƬ����Ϊ�գ�");return false;}
if(npcdiren=="" ){alert("��ʾ��NPc���˲���Ϊ�գ�");return false;}
if(npczhuren=="" ){alert("��ʾ��NPc���˲���Ϊ�գ�");return false;}
if(npcwp=="" ){alert("��ʾ��NPc��Ʒ����Ϊ�գ�");return false;}
if(npczs=="" ){alert("��ʾ��NPc��ʽ����Ϊ�գ�");return false;}
if(npcccc=="" ){alert("��ʾ��NPc�����ʲ���Ϊ�գ�");return false;}
if(npcwplx=="" ){alert("��ʾ��NPc��Ʒ���Ͳ���Ϊ�գ�");return false;}
var pattern = /^([0-9])+$/;
if(pattern.test(npcvalue)!=true){alert("��ʾ��npc������ʹ�����֣�");return false;}
if(pattern.test(npcmoney)!=true){alert("��ʾ��npc������ʹ�����֣�");return false;}
if(pattern.test(npcgj)!=true){alert("��ʾ��npc������ʹ�����֣�");return false;}
if(pattern.test(npcfy)!=true){alert("��ʾ��npc������ʹ�����֣�");return false;}
if(pattern.test(npcwg)!=true){alert("��ʾ��npc�书��ʹ�����֣�");return false;}
if(pattern.test(npctl)!=true){alert("��ʾ��npc������ʹ�����֣�");return false;}
if(pattern.test(npcnl)!=true){alert("��ʾ��npc������ʹ�����֣�");return false;}
if(pattern.test(npcmf)!=true){alert("��ʾ��npcħ����ʹ�����֣�");return false;}
if(pattern.test(npcgjl)!=true){alert("��ʾ��npc��������ʹ�����֣�");return false;}
if(pattern.test(npccxl)!=true){alert("��ʾ��npc��������ʹ�����֣�");return false;}
if(pattern.test(npcdie)!=true){alert("��ʾ��npc��������ʹ�����֣�");return false;}
}
</script>
<form name="form" method="post" action="jhgl.asp?act=modinpc&id=<%=rs("id")%>" onSubmit="return check(this);">
  <div align="center"></div>
  <table width="85%" border=1 align="center" cellpadding=1 cellspacing=0  bordercolor="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor="#d7ebff"> 
      <td colspan="10" align="center">�� �� NPC (<font color="#FF0000"><strong>���밴��ʽ���,��������������</strong></font>)</td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">NPC����</td>
      <td> <input name="npcname" id="npcname" value="<%=rs("n����")%>" size=10 maxlength=10> 
      </td>
      <td align="center" bgcolor="#d7ebff">�Ա�</td>
      <td> <select name="npcsex" id="npcsex">
          <option value="��" <%if trim(rs("n�Ա�"))="��" then%>selected<%end if%>>��</option>
          <option value="Ů" <%if trim(rs("n�Ա�"))="Ů" then%>selected<%end if%>>Ů</option>
        </select></td>
      <td align="center" bgcolor="#d7ebff">����</td>
      <td> <input name="npcvalue" id="npcvalue" value="<%=rs("n����")%>" size=10 maxlength=10> 
      </td>
      <td align="center" bgcolor="#d7ebff">ͼƬ</td>
      <td><input name="npcimg" type="text" id="npcimg" value="<%=trim(rs("nͼƬ"))%>" size="10" maxlength="50"></td>
      <td align="center" bgcolor="#d7ebff">����</td>
      <td><input name="npcmoney" type="text" id="npcmoney" value="<%=rs("n����")%>" size="10" maxlength="15"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">����</td>
      <td> <input name="npcgj" id="npcgj" value="<%=rs("n����")%>" size=10 maxlength=10> </td>
      <td align="center" bgcolor="#d7ebff">����</td>
      <td> <input name="npcfy" type="text" id="npcfy" value="<%=rs("n����")%>" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">�书</td>
      <td> <input name="npcwg" id="npcwg" value="<%=rs("n�书")%>" size=10 maxlength=10> </td>
      <td align="center" bgcolor="#d7ebff">����</td>
      <td><input name="npctl" type="text" id="npctl" value="<%=rs("n����")%>" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">����</td>
      <td><input name="npcnl" type="text" id="npcnl" value="<%=rs("n����")%>" size="10" maxlength="10"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">������</td>
      <td><input name="npcdie" type="text" id="npcdie" value="<%=rs("n������")%>" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">������</td>
      <td><input name="npcgjl" type="text" id="npcgjl" value="<%=rs("n������")%>" size="10" maxlength="10"> 
      </td>
      <td align="center" bgcolor="#d7ebff">������</td>
      <td><input name="npccxl" type="text" id="npccxl" value="<%=rs("n������")%>" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">����</td>
      <td><input name="npczhuren" type="text" id="npczhuren" value="<%=rs("n����")%>" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">����</td>
      <td><input name="npcdiren" type="text" id="npcdiren" value="<%=rs("n����")%>" size="10" maxlength="10"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">��Ʒ</td>
      <td colspan="7"><input name="npcwp" type="text" id="npcwp" value="<%=trim(rs("n��Ʒ"))%>" size="60" maxlength="200"></td> 
      <td align="center" bgcolor="#d7ebff">��Ʒ����</td>
      <td><input name="npcwplx" type="text" id="npcwplx" value="<%=trim(rs("n��Ʒ����"))%>" size="4" maxlength="3"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">������</td>
      <td colspan="9"><input name="npcccc" type="text" id="npcccc" value="<%=trim(rs("n������"))%>" size="70" maxlength="200">
        <input name="npcid" type="hidden" id="npcid" value="<%=rs("id")%>"></td>
    </tr>
    <tr> 
      <td colspan="10" align="center" valign=top bgcolor="#d7ebff"> <input type="SUBMIT" name="Submit" value="�޸�"> 
        <input type="RESET" name="Reset" value="���"> </td>
    </tr>
  </table>
  <br>
  <table width="85%" border="0" align="center">
    <tr>
      <td><font color="#0000FF"><strong>����:</strong></font>������ô���NPC�ĵȼ�,������ֵ����:�ȼ�x�ȼ�x50=������,��һ��Ҳ��Ӱ�쵽���Ĺ���ֵ.<br>
        <font color="#0000FF"><strong><br>
        ͼƬ:</strong></font>���Ҫָ����ͼƬ���ڵ�λ��,���ﲻ��������������,��Ҫʹ��;��'��2������.<br>
        <strong><font color="#0000FF"><br>
        ������:</font></strong>rnd&lt;1/������,��Ϊ1ʱnpc����Ϊ100%,ֵԽ�󹥻������ԽС.������,���д��жϱ�������npc����ʱ����.<br>
        <br>
        <font color="#0000FF"><strong>������:</strong></font>����ʱ���뵱ǰ���ڵĲ�������ڴ�ֵ,���npc���Ա�����ϵͳ��.ԽС,npc�ֳ�����Խ��.��λ����.<br>
        <br>
        <font color="#0000FF"><strong>��Ʒ:</strong></font> npc�������Ʒ,��ʽΪ:������1|��Ʒ��1,������2|��Ʒ��2,Ҫ˵����������ĳ�����,����ֵ:<br>
        (������/������)=ȡ��(������/������) �����ͬ,���������,�ȷ����ó�20,ֻ��NPCÿ��20�βŻ�����Ʒ����,���ó�1������һ�ζ������.<br>
        <br>
        <strong><font color="#0000FF">��ʽ:</font></strong>Npc����ʱʹ�õ���ʽ,�����������ʱ,����ʹ�û�����ʽ���й���,����ҧ,ץ,�ߵ�.��3��������ʽ�����ݿ��и���Ҫɾ��.�����޷�ʹ��.<br>
        <br>
        <font color="#0000FF"><strong>������:</strong></font>Npc�ֳ�ʱ�Ļ�������,���ﲻҪ����&quot;��'��.��������ע��.<br>
        <br>
        
      </td>
    </tr>
  </table>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</form>
<br>
</body>
</html>
