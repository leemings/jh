<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=20 background="../chatroom/bg1.gif">
<div align=center>
�����ѯ��ѯ<form method="POST" action="showuser2.asp">
  <div align="center"> 
    <select name="sjcz">
      <option value="����" selected>����</option>
      <option value="ID">ID</option>
      <option value="��ż">��ż</option>
      <option value="��������">����</option>
      <option value="�ʺ�">�ʺ�</option>
      <option value="Oicq">OIcq</option>
    </select>
    <input type="text" name="search" size="15" maxlength="20">     
    <input type="submit" value="��ѯ" name="B1" class="p9">     
    <input type="reset" name="Submit" value="���">     
  </div>     
  <div align="center">ID����һ��ҪΪ���֣�<br>     
  </div>     
  </form>    
��IP��ѯ<form action="showbyip.asp" method=post><input type=text name=ip maxlength=15 size=15 value=''><input type=submit value=' �� �� '></form>   
���ȼ���ѯ<form action=showbygr.asp method=post>   
<select name=grade>   
<option value=-5>-5��</option>
<option value=-4>-4��</option>
<option value=-3>-3��</option>
<option value=-2>-2��</option>
<option value=-1>-1��</option>
<option value=0>0��</option>
<option value=1 selected>1��</option>
<option value=2>2��</option>
<option value=3>3��</option>
<option value=4>4��</option>
<option value=5>5��</option>
<option value=6>6��</option>
<option value=7>7��</option>
<option value=8>8��</option>
<option value=9>9��</option>
<option value=10>10��</option>
<option value=11>11��</option>
<option value=12>12��</option>
<option value=13>13��</option>
<option value=14>14��</option>
<option value=15>15��</option>
<option value=16>16��</option>
<option value=17>17��</option>
<option value=18>18��</option>
<option value=19>19��</option>
<option value=20>20��</option>
<option value=21>21��</option>
<option value=21>21��</option>
<option value=23>23��</option>
<option value=24>24��</option>
<option value=25>25��</option>
<option value=26>26��</option>
<option value=27>27��</option>
<option value=28>28��</option>
<option value=29>29��</option>
<option value=30>30��</option>
<option value=31>31��</option>
<option value=32>32��</option>
<option value=33>33��</option>
<option value=34>34��</option>
<option value=35>35��</option>
<option value=36>36��</option>
<option value=37>37��</option>
<option value=34>38��</option>
<option value=35>39��</option>
<option value=36>40��</option>
<option value=37>41��</option>
</select><input type=submit value='����'></form>
�����ɲ�ѯ<form action=showbyco.asp method=post><select name=corp>
<option value='��' selected>��������</option>
<%
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "����",conn
do while not rst.EOF
Response.Write "<option value='"&rst("����")&"'>"&rst("����")&"</option>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
</select><input type=submit value='����'></form>
ɾ�������ʺ�<form action=deluser.asp method=post>
<select name=howday>
<option value=7>7��</option>
<option value=15>15��</option>
<option value=30 selected>30��</option>
<option value=60>60��</option>
<option value=90>90��</option>
</select>
<input type=submit value='ȷ��ɾ��'>   
</form>   
    
<form method="POST" action="laren.asp">     
  <div align="center">���˲鿴����<br>     
    <input type="text" name="uname" size="10" maxlength="10">     
    <input type="submit" value="��ѯ" name="B12" class="p9">     
    <input type="reset" name="Submit2" value="���">      
  <div align="center"></div>     
</form>     
<form method="POST" action="password.asp">     
  <div align="center">������������<br>������������������ĳɣ�123456<br>     
    <input type="text" name="cpass" size="10" maxlength="10">     
    <input type="submit" value="�޸�" name="B12" class="p9">     
    <input type="reset" name="Submit2" value="���">     
    <br>     
  </div>     
  </form>     
</body>





