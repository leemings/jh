<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>��Ա���ݿ�����wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><%=Application("sjjh_chatroomname")%>���ݿ�</p>
<table width="572" border="1" align="center" cellspacing="0" cellpadding="0" bordercolor="#000000">
<tr>
<td>
      <p align="center"><a href="hygl.asp">��Ա�����������</a></p>
<form method="POST" action="hydjok.asp">
        <div align="center">
          <p><b><font color="#FF0000">ע�⣺</font></b><font color="#FF0000">�����Ѿ��ǻ�Ա�ģ�������ɺ�ʱ��ᰴ��Ա����ʱ���<br>
            ������ʱ�䣬��Ա�ȼ������ó��µĵȼ�!</font></p>
          <p><font color="#FF0000"><br>
            </font>��Ա����ʱ�䣺 
            <select name=hymonth   size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
              <option value="1" selected>01</option>
              <option value="2">02</option>
              <option value="3">03</option>
              <option value="4">04</option>
              <option value="5">05</option>
              <option value="6">06</option>
              <option value="7">07</option>
              <option value="8">08</option>
              <option value="9">09</option>
              <option value="10">10</option>
              <option value="11">11</option>
              <option value="12">12</option>
            </select>
            �� <br>
            ��Ա�ȼ��� 
            <select name="hygrade">
              <option value="1" selected>һ����Ա</option>
              <option value="2">������Ա</option>
              <option value="3">������Ա</option>
              <option value="4">�ļ���Ա</option>
            </select>
            <br>
            ��Ա������ 
            <input type="text" name="search" size="10" maxlength="10">
            <br>
            <input type="submit" value="ȷ���»�Ա" name="B1" class="p9">
          </p>
        </div>
      </form>
<div align="center"><font color="#FF0000"><br>
          ǧ��Ҫ��վ�����û�Ա�������ƻ�Ա���ֶ����ã�</font><br>
<br>
</div>
<div align="center">��ѡ����Ҫ���������ȼ����ã�<br>
��Ҫע�⣺��Ա��ֹ���ڲ���С�ڵ�ǰ���ڣ���Ա�ڵ���ǰ5����ʾ��<br>
��ͬ�ȼ��Ļ�Ա�ӵ�Ҳ��ͬ�ģ�1����Ա��ս��25 2����Ա��ս��40 ���޼ӣ�100<br>
3����Ա:ս��70�� ���޼ӣ�200 4����ս��98 ����:7�� ���ޣ�300 ���Ҫ�޸����ֶ��ģ�<br>
</div>
<p align="center">[վ������ͨ���˴����޸��û�����������]</p>
<p align="center"><a href="../jhmp/hy.asp?fs=1">���л�Ա�б�</a> <a href="fine.asp">�����û�����</a></p>
</td>
</tr>
</table>
<p align="center">&nbsp;</p>
</body>
</html>
