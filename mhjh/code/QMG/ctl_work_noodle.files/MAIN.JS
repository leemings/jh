function validate_search()
{
	if (document.searchform.searchtext.value == "")
	{
	alert("������Ҫ���ҵ�����");
	return false;
	}
	if (document.searchform.scope.value == "user")
	{
	//document.searchform.action = "/resident/ctl_resident.php?action=search";
	document.searchform.action = "/resident/search_pop.php";
	document.searchform.name.value = document.searchform.searchtext.value;
	return true;
	}
	else
	{
		if (document.searchform.scope.value == "download")
		{
		//	Added Dennis Zou
		document.searchform.action = "http://www.the9.com/download/dl_filelist.php";
		document.searchform.dl_key_words.value = document.searchform.searchtext.value;
		return true;
		}
		else
		{
		alert("Ŀǰ��֧�����ֲ��ҷ�ʽ");
		return false;
		}

	}
}

function validate()
{
	if (document.loginform.loginname.value == "")
	{
		alert("�������¼��");
		document.loginform.loginname.focus();
		return;
	}
	if (document.loginform.password.value == "")
	{
		alert("���������");
		document.loginform.password.focus();
		return;
	}
	document.loginform.submit();
}

function go_bbs(sel)
{
	var select_id=sel.selectedIndex;
	if (select_id==0)
	{
		alert("��ѡ����̳");
		return false;
	}
	else
	{
		var bbs_url="http://bbs.the9.com/bbs/index.php?forum=" + sel.options[select_id].value;
		window.location=bbs_url;
		return true;
	}
	
}

function keysubmit(e)
{
	if (navigator.appName == "Netscape")
	{
		if (e.which == 13){validate();}
	}
	else
	{
		if (event.keyCode == 13){validate();}
	}
}

function getpass()
{
	window.open("http://www.the9.com/indexhelp/get_password.htm", 'newwin2', 'left=150, top=50, width=600,height=300,resizable=1,status=0,menubar=0,scrollbars=1'); 
}
 
