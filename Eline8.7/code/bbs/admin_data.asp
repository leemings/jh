<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
dim action
dim admin_flag
action=trim(request("action"))

dim dbpath,bkfolder,bkdbname,fso,fso1

select case action
case "CompressData"		'ѹ������
	admin_flag="61"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		call CompressData()

	end if	

case "BackupData"		'��������
	admin_flag="62"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("act")="Backup" then
		call updata()
		else
		call BackupData()
		end if
	end if

case "RestoreData"		'�ָ�����
	admin_flag="63"
	dim backpath
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
			response.write "��������Ҫ�ָ��ɵ����ݿ�ȫ��"	
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
		
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			response.write "�ɹ��ָ����ݣ�"
			else
			response.write "����Ŀ¼�²������ı����ļ���"	
			end if
		else
		call RestoreData()
		end if
	end if

 case "SpaceSize"		'ϵͳ�ռ�ռ��
	admin_flag="64"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else	
		call SpaceSize()
	end if

case else
		Errmsg=Errmsg+"<br>"+"<li>ѡȡ��Ӧ�Ĳ�����"
		call dvbbs_error()

end select

conn.close
set conn=nothing
response.write"</body></html>"


'====================ϵͳ�ռ�ռ��=======================
sub SpaceSize()
on error resume next
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;ϵͳ�ռ�ռ�����
  					</th>
  				</tr> 	
 				<tr>
 					<td class="forumrow"> 			
 			<blockquote>
 			<br> 			
 			��������ռ�ÿռ䣺&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("data")%> height=10>&nbsp;<%showSpaceinfo("data")%><br><br>
 			��������ռ�ÿռ䣺&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("databackup")%> height=10>&nbsp;<%showSpaceinfo("databackup")%><br><br>
 			�����ļ�ռ�ÿռ䣺&nbsp;<img src="pic/bar1.gif" width=<%=drawspecialbar%> height=10>&nbsp;<%showSpecialSpaceinfo("Program")%><br><br>
 			����ͼƬռ�ÿռ䣺&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("images")%> height=10>&nbsp;<%showSpaceinfo("face")%><br><br>
 			ϵͳͼƬռ�ÿռ䣺&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("pic")%> height=10>&nbsp;<%showSpaceinfo("pic")%><br><br>
 			�ϴ�ͷ��ռ�ÿռ䣺&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("uploadFace")%> height=10>&nbsp;<%showSpaceinfo("uploadFace")%><br><br>
 			�ϴ��ļ�ռ�ÿռ䣺&nbsp;<img src="pic/bar1.gif" width=<%=drawbar("uploadFile")%> height=10>&nbsp;<%showSpaceinfo("uploadFile")%><br><br>	
 			ϵͳռ�ÿռ��ܼƣ�<br><img src="pic/bar1.gif" width=400 height=10> <%showspecialspaceinfo("All")%>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
end sub

'====================�ָ����ݿ�=========================
sub RestoreData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder"		<tr>
	<th height=25 >
   					&nbsp;&nbsp;<B>�ָ���̳����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
				<form method="post" action="ADMIN_data.asp?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;�������ݿ�·��(���)��<input type=text size=30 name=DBpath value="DataBackup\eline_bbs_6.3.0.asp">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;Ŀ�����ݿ�·��(���)��<input type=text size=30 name=backpath value="<%=db%>"><BR>&nbsp;&nbsp;��д����ǰʹ�õ����ݿ�·�����粻�븲�ǵ�ǰ�ļ���������������ע��·���Ƿ���ȷ����Ȼ���޸�conn.asp�ļ������Ŀ���ļ����͵�ǰʹ�����ݿ���һ�µĻ��������޸�conn.asp�ļ�<BR>
						&nbsp;&nbsp;<input type=submit value="�ָ�����"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�ϱ������ݿ��ļ�ΪDataBackup\DVBBS60.MDB���밴�����ı����ļ������޸ġ�<br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

'====================�������ݿ�=========================
sub BackupData()
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>������̳����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_data.asp?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;
						��ǰ���ݿ�·��(���·��)��<input type=text size=15 name=DBpath value="<%=db%>"><BR>&nbsp;&nbsp;
						�������ݿ�Ŀ¼(���·��)��<input type=text size=15 name=bkfolder value=Databackup>&nbsp;��Ŀ¼�����ڣ������Զ�����<BR>&nbsp;&nbsp;
						�������ݿ�����(��д����)��<input type=text size=15 name=bkDBname value=DVBBS60.MDB>&nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����<BR>
						&nbsp;&nbsp;<input type=submit value="ȷ��"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�����ݿ��ļ�ΪData\DVBBS60.MDB��<B>��һ��������Ĭ�����������������ݿ�</B><br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��				</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

sub updata()
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			response.write "�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname
		Else
			response.write "�Ҳ���������Ҫ���ݵ��ļ���"
		End if
end sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================ѹ�����ݿ� =========================
sub CompressData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<form action="Admin_data.asp?action=CompressData" method="post">
<tr>
<td class="forumrow" height=25><b>ע�⣺</b><br>�������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ�������� </td>
</tr>
<tr>
<td class="forumrow">ѹ�����ݿ⣺<input type="text" name="dbpath" value=Data\DVBBS60.MDB>&nbsp;
<input type="submit" value="��ʼѹ��"></td>
</tr>
<tr>
<td class="forumrow"><input type="checkbox" name="boolIs97" value="True">���ʹ�� Access 97 ���ݿ���ѡ��
(Ĭ��Ϊ Access 2000 ���ݿ�)<br><br></td>
</tr>
<form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================ѹ������=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "������ݿ�, " & dbpath & ", �Ѿ�ѹ���ɹ�!" & vbCrLf

Else
	CompactDB = "���ݿ����ƻ�·������ȷ. ������!" & vbCrLf
End If

End Function


'=====================ϵͳ�ռ����=========================
	Sub ShowSpaceInfo(drvpath)
 		dim fso,d,size,showsize
 		set fso=server.createobject("scripting.filesystemobject") 		
 		drvpath=server.mappath(drvpath) 		 		
 		set d=fso.getfolder(drvpath) 		
 		size=d.size
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	End Sub	
 	
 	Sub Showspecialspaceinfo(method)
 		dim fso,d,fc,f1,size,showsize,drvpath 		
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpath=server.mappath("pic")
 		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
 		set d=fso.getfolder(drvpath) 		
 		
 		if method="All" then 		
 			size=d.size
 		elseif method="Program" then
 			set fc=d.Files
 			for each f1 in fc
 				size=size+f1.size
 			next	
 		end if	
 		
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	end sub 	 	 	
 	
 	Function Drawbar(drvpath)
 		dim fso,drvpathroot,d,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		drvpath=server.mappath(drvpath) 		
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		
 		barsize=cint((size/totalsize)*400)
 		Drawbar=barsize
 	End Function 	
 	
 	Function Drawspecialbar()
 		dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		set fc=d.files
 		for each f1 in fc
 			size=size+f1.size
 		next	
 		
 		barsize=cint((size/totalsize)*400)
 		Drawspecialbar=barsize
 	End Function 	
%>