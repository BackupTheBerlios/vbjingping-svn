<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->

<%sub register()
name1=REQUEST.FORM("name")
code=REQUEST.FORM("password")
sex=REQUEST.FORM("sex")
birth=REQUEST.FORM("birth")
com=REQUEST.FORM("com")
if name1="" or code="" then
%>
<table border="0" cellpadding="10" cellspacing="0" width=90% align=center><tr><td>
<br>
<FORM METHOD="post" ACTION="register.asp" ID="Form1">
           <b>ע��:</b><br>
           ��&nbsp;��&nbsp;����<INPUT TYPE="text" name="name" SIZE="22" ID="Text1"><br>
           ��&nbsp;&nbsp;&nbsp;&nbsp;�룺<INPUT TYPE="password" name="password" SIZE="28" ID="Password1" style="font-size:12px"><br>
           ��&nbsp;&nbsp;&nbsp;&nbsp;��<input type="radio" name="sex" value="��" checked="checked" />��
           <input type="radio" name="sex" value="Ů" />Ů
           <br>
           �������£�<INPUT TYPE="text" name="birth" SIZE="22" ID="Text3"><br>
           ��ϵ��ʽ��<INPUT TYPE="text" name="com" SIZE="22"><br>
           <INPUT TYPE="submit" VALUE="����ע��" id="regcss"> 
           <INPUT TYPE="reset" VALUE="��������" id="reregcss">
</FORM> 

</td></tr></table>
<%else
rsStr="select * from author where name = " & "'" & name1 & "'"
set rs=conn.execute(rsStr)
If Not rs.EOF Then %> 
           <!-- ������û����Ѵ��ڣ������䣬����������ݿ� -->
           <br>
           ���û����ѱ�ע�ᣬ��������<a href="register.asp">ע��</a>���û���! 
<% Else
           rs.Close
           set rs=conn.execute("insert into author(name,code,sex,birth,com) values('" _
           &name1& "','" &code& "','" &sex& "','" &birth& "','" &com&"')")
           conn.Close
           set rs=nothing
%>
           <CENTER>
           <br>
           <B><% =name1 %></B> ����ע��ɹ�!<P>
           <a href="index.asp">������̳</a>|<a href="say.asp?name=<%=name1%>">��������</a>
           </CENTER>
<% End If %>
<br> 
<%end if%>
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
<title>ע��</title>
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
		<table width="770" border="0" cellspacing="0" cellpadding="0" align="center">
               <tr>
                 <td id="tbul">&nbsp;</td>
                 <td id= "tbum"><%headtext()%></td>
                 <td id="tbur">&nbsp;</td>
               </tr>
               <tr>
                 <td id="tbml">&nbsp;</td>
                 <td id="tbmm"><%register()%></td>
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
>

