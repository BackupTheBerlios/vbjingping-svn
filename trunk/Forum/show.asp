<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->
<%id=request.querystring("id")
set show=conn.execute("select*from bbs where id="&id&"")'ȫ��%>
<%sub maintalk
'������%>
<br>
<div id="arttitle">
<%=show("title")%><br><br>
</div>
<div id="authorn">
����:<a href="authordetail.asp?auname=<%=show("name")%>"><%=show("name")%></a>&nbsp;
׫д/�޸�ʱ��:<%=show("wtime")%>&nbsp;
<a href="modify.asp?id=<%=show("id")%>">�޸�</a>
<a href="del.asp?id=<%=show("id")%>">ɾ��</a>
</div>
<table border="0" id="content_table"><tr><td>
<div id="content"><b>���ݣ�</b><%=show("body")%></div>
</td></tr></table>
<br><hr size="1" width=700 align="left">
<%end sub%>
<%sub wbtalk()
'�ظ�����%>
<%sql="select*from bbbs where hostid='" & id & "' order by btime asc"
set showwback=conn.execute(sql)
floori=0
do while not showwback.eof 'û������ҳ������ҳ��ԭ����ͬ��
floori=floori+1%>
<div id="arttitle">
<%=floori%>¥:<%=showwback("btitle")%><br><br>
</div>
<div id="authorn">
�ظ��ˣ�<a href="authordetail.asp?auname=<%=showwback("bname")%>"><%=showwback("bname")%></a>
�ظ�ʱ��:<%=showwback("btime")%>
</div>
<table border="0" id="content_table"><tr><td>
<div id="content"><b>�ظ����ݣ�</b><%=showwback("bcontent")%></div>
</td></tr></table>
<br>
<!--<hr size="1" width=600 align=left style=" border:dashed #999999 "> -->
<%showwback.movenext 
Loop
set showwback=nothing
end sub%>
<%sub wbform()%>
<br><hr size="1" width=700 align=center>
      <table border="0" cellpadding="0" cellspacing="0" 
       style="border-collapse: collapse; " bordercolor="#000000" 
       width=600 height="20" align=center ID="Table1">
<tr><td>
<form method="POST" action="wback.asp" >
��&nbsp;��&nbsp;�ˣ�<input type="text" name="name_b" size="20"> 
����:<input type="password" name="code_b" size="20" style="font-size:12px">
���<a href="register.asp">ע��</a><br>
<input type="hidden" name="id_h" size="20"      value= <%=id%> />
�ظ����⣺<input type="text" name="title_b" size="70" value="re:<%=show("title")%>"><br>
�ظ����ݣ�<br>
<textarea rows="11" name="body_b" cols="80"></textarea><br> 
<br>
<input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"> 
</form>
</td></tr></table>
<%end sub%>
<html><head>
<link href="main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
	.top{
	vertical-align:top;}
	.btn{
	background-color:#CCFFCC;
	width:130px;
	height:30px;
	}

</style>
<title>�鿴����</title>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (sub01��.psd) -->
<table id="__01" width="950" height="793" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="3">
			<img src="images/sub01��_01.gif" width="814" height="90" alt=""></td>
		<td colspan="2">
			<img src="images/sub01��_02.gif" width="136" height="90" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/sub01_03.gif" width="950" height="68" alt=""></td>
	</tr>
	<tr>
		<td><a href="index.asp"><img src="images/left_menu01.gif" width="219" height="29" alt="" border="0"></a></td>
		<td colspan="4" rowspan="2">
			<img src="images/sub01��_05.gif" width="731" height="49" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/left_menu02.gif" width="219" height="20" alt=""></td>
	</tr>
	<tr>
		<td rowspan="2" class="top">
			<img src="images/sub01��_07.gif" width="219" height="505" alt=""></td>
		<td rowspan="2">
			<img src="images/sub01��_08.gif" width="21" height="505" alt=""></td>
		<td colspan="2" background="images/txt.gif">
		<table width="770" border="0" cellspacing="0" align="center">
      <tr>
        <td id="tbul">&nbsp;</td>
        <td id= "tbum"><%headtext()%></td>
        <td id="tbur">&nbsp;</td>
      </tr>
      <tr>
        <td id="tbml">&nbsp;</td>
        <td id="tbmm"><%maintalk()%><%wbtalk()%><%wbform()%></td>
        <td id="tbmr">&nbsp;</td>
      </tr>
  
      <tr>
        <td id="tbll">&nbsp;</td>
        <td id="tblm">&nbsp;</td>
        <td id="tblr">&nbsp;</td>
      </tr>
</table>
<!--#include file="footer.htm"--></td>
		<td rowspan="2">
			<img src="images/sub01��_10.gif" width="18" height="505" alt=""></td>
	</tr>
	<tr>
		<td colspan="2">
			<img src="images/sub01��_11.gif" width="692" height="14" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/sub01��_12.gif" width="950" height="80" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/�ָ���.gif" width="219" height="1" alt=""></td>
		<td>
			<img src="images/�ָ���.gif" width="21" height="1" alt=""></td>
		<td>
			<img src="images/�ָ���.gif" width="574" height="1" alt=""></td>
		<td>
			<img src="images/�ָ���.gif" width="118" height="1" alt=""></td>
		<td>
			<img src="images/�ָ���.gif" width="18" height="1" alt=""></td>
	</tr>
</table>
<!-- End ImageReady Slices -->
</body>
</html>

