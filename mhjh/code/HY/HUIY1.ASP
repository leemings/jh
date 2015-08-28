<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 电子邮箱,签名档 from 用户 where 姓名='"&username&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF  then Response.Redirect "error.asp?id=019"
email=rst("电子邮箱")
qianming=rst("签名档")
%>
<head><title>会员修改</title><link rel="stylesheet" href="../style.css"></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false >
<p align="center"><font size="3">会员修改</font></p>
<form action=huiy2.asp method=post>
<table border=1 width=677 align=center height="68" bordercolorlight="#000000" cellspacing="0" bordercolordark="#FFFFFF">
<tr><td height="23" width="155"><font color="#000000">电子邮箱</font></td><td height="23" width="502"><INPUT class=norsty maxLength=30 name=e_Mail size=30 value=<%=email %>></td></tr><tr> 
            <td height="34" width="155"> 头像<img id="face" src="../face/16.gif"> </td>
            <td height="34" width="502"> 
             <select name="faceimg"  size=1 onChange="document.images['face'].src=options[selectedIndex].value;"> 
                <option value="../face/16.gif"> 男001</option>
                <option value="../face/NAN002.gif"> 男002</option>
                <option value="../face/NAN003.gif"> 男003</option>
                <option value="../face/NAN004.gif"> 男004</option>
                <option value="../face/NAN005.gif"> 男005</option>
                <option value="../face/NAN006.gif"> 男006</option>
                <option value="../face/NAN007.gif"> 男007</option>
                <option value="../face/NAN008.gif"> 男008</option>
                <option value="../face/NAN009.gif"> 男009</option>
                <option value="../face/NAN010.gif"> 男010</option>
                <option value="../face/NAN011.gif"> 男011</option>
                <option value="../face/NAN012.gif"> 男012</option>
                <option value="../face/NAN013.gif"> 男013</option>
                <option value="../face/NAN014.gif"> 男014</option>
                <option value="../face/NAN015.gif"> 男015</option>
                <option value="../face/NAN016.gif"> 男016</option>
                <option value="../face/NAN017.gif"> 男017</option>
                <option value="../face/NAN018.gif"> 男018</option>
                <option value="../face/NAN019.gif"> 男019</option>
                <option value="../face/NAN020.gif"> 男020</option>
                <option value='../face/NAN021.gif'> 男021</option>
                <option value='../face/NAN022.gif'> 男022</option>
                <option value='../face/NAN023.gif'> 男023</option>
                <option value='../face/NAN024.gif'> 男024</option>
                <option value='../face/NAN025.gif'> 男025</option>
                <option value='../face/NAN026.gif'> 男026</option>
                <option value='../face/NAN027.gif'> 男027</option>
                <option value='../face/NAN028.gif'> 男028</option>
		        <option value='../face/NAN001.gif'> 男029</option>
                <option value='../face/NAN030.gif'> 男030</option>
		        <option value='../face/NAN031.gif'> 男031</option>
		        <option value='../face/NAN032.gif'> 男032</option>
                <option value='../face/NV001.gif'> 女001</option>
                <option value='../face/NV002.gif'> 女002</option>
                <option value='../face/NV003.gif'> 女003</option>
                <option value='../face/NV004.gif'> 女004</option>
                <option value='../face/NV005.gif'> 女005</option>
                <option value='../face/NV006.gif'> 女006</option>
                <option value='../face/NV007.gif'> 女007</option>
                <option value='../face/NV008.gif'> 女008</option>
                <option value='../face/NV009.gif'> 女009</option>
                <option value='../face/NV010.gif'> 女010</option>
                <option value='../face/NV011.gif'> 女011</option>
                <option value='../face/NV012.gif'> 女012</option>
                <option value='../face/NV013.gif'> 女013</option>
                <option value='../face/NV014.gif'> 女014</option>
                <option value='../face/NV015.gif'> 女015</option>
                <option value='../face/NV016.gif'> 女016</option>
                <option value='../face/NV017.gif'> 女017</option>
                <option value='../face/NV018.gif'> 女018</option>
                <option value='../face/NV019.gif'> 女019</option>
                <option value='../face/NV020.gif'> 女020</option>
                <option value='../face/NV021.gif'> 女021</option>
                <option value='../face/NV022.gif'> 女022</option>
                <option value='../face/NV023.gif'> 女023</option>
                <option value='../face/NV024.gif'> 女024</option>
                <option value='../face/NV025.gif'> 女025</option>
                <option value='../face/NV026.gif'> 女026</option>
                <option value='../face/NV027.gif'> 女027</option>
                <option value='../face/NV028.gif'> 女028</option>
                <option value='../face/NV029.gif'> 女029</option>
                <option value='../face/NV030.gif'> 女030</option>
                <option value='../face/NV031.gif'> 女031</option>
              </select>
            </td>
          </tr><tr><td height="1" width="155"><font color="#000000">签名档</font></td><td height="1" width="502"><textarea name="qianming" cols="30"  rows="5" style="font-family: Tahoma; font-size: 12px"><%=qianming%></textarea></td></tr>
<tr align=center><td colspan=2 height="25" width="663"><input type=submit value=" 更 改 "> <input type=button onclick=javascript:history.back(); value=' 返 回 '></td></tr></table>
</form>
</body>
<%rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>