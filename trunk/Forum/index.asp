<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->
<%sub maintalk()%>
<%
set showbbs = server.CreateObject("ADODB.recordset")
showbbsStr="select*from bbs order by wtime desc"
showbbs.open showbbsStr,conn,3,2
if showbbs.EOF and showbbs.BOF then%>
<br><center>��ʱ��û������,����<a href="say.asp">����</a></center>
<%else
showbbs.PageSize=10'��ҳ
PageN=10 '��ʾ10��ҳ��.
PageCount=showbbs.PageCount
Page=int(request("Page"))
CurrentPageN=int(request("CurrentPageN"))
if Page<=0 or request("Page")="" or request("Page")="0" then Page=1
if CurrentPageN<=0 or request("CurrentPageN")="" then CurrentPageN=1
showbbs.AbsolutePage=Page
%>
<br>
<%for i=1 to showbbs.PageSize%>
<div id="arttitle">
      <a href="show.asp?id=<%=showbbs("id")%>"><%=showbbs("title")%></a><br><br>
</div>

<div id="authorn">
      ����:<a href="authordetail.asp?auname=<%=showbbs("name")%>"><%=showbbs("name")%></a>&nbsp;
      <%if showbbs("countwb")=0 then%>
       �ظ�&lt;<%=showbbs("countwb")%>&gt;<!--�ظ����� -->
      <%else%>
       <a href="show.asp?id=<%=showbbs("id")%>">
       �ظ�&lt;<font color=#ff0000><b><%=showbbs("countwb")%></b></font>&gt;
       </a><!--�ظ����� -->
      <%end if%>&nbsp;
  
      <%wt=showbbs("wtime")'���Ϊ׫д�����ʱ�䣬��ɫ��ʾ��
      if year(now())=year(wt) and month(now())=month(wt) and day(now())=day(wt) then%>
       ׫д/�޸�ʱ��:<font color=#ff0000><%=wt%></font>&nbsp;
      <%else%>
       ׫д/�޸�ʱ��:<%=wt%>&nbsp;
      <%end if%>
  
      <a href="modify.asp?id=<%=showbbs("id")%>">�޸�</a>
      <a href="del.asp?id=<%=showbbs("id")%>">ɾ��</a>
</div>

<hr size="1" width=600  align="left" style=" border:dashed #999999 ">

<%showbbs.movenext
if showbbs.EOF then exit for
Next%>
<center>
<br>
<%
if Page<>1 then%>
<a href="index.asp?Page=1&CurrentPageN=1">(��ҳ)</a>
<%else%>
(��ǰΪ��ҳ)
<%end if%>
<%if PageCount>PageN and CurrentPageN>=2 then%>
<a href="index.asp?Page=<%=(CurrentPageN-2)*PageN+1%>&CurrentPageN=<%=CurrentPageN-1%>">(��<%=PageN%>ҳ)</a>
<%end if%>
<%for i=(CurrentPageN-1)*PageN+1 to (CurrentPageN-1)*PageN+PageN%>
<%if i = Page then%>
      <font color=#ff0000 >-<%=i%>-</font>
<%elseif i<showbbs.RecordCount/showbbs.PageSize +1 then%>
      <a href="index.asp?page=<%=i%>&CurrentPageN=<%=CurrentPageN%>">&nbsp;<%=i%>&nbsp;</a>
<%end if%>
<%Next%>
<%if PageCount>PageN and CurrentPageN<PageCount/PageN then%>
<a href="index.asp?Page=<%=CurrentPageN*PageN+1%>&CurrentPageN=<%=CurrentPageN+1%>">(��<%=PageN%>ҳ)</a>
<%end if%>
<%if PageCount>PageN and Page<>PageCount then%>
<a href="index.asp?Page=<%=PageCount%>&CurrentPageN=<%=int(PageCount/PageN)+1%>">(���һҳ)</a>
<%end if%>
<%if PageCount=Page then%>(��ǰΪ���ҳ)<%end if%>
(��<%=PageCount%>ҳ)
</center>
<%showbbs.Close
set showbbs=nothing
end if
end sub%>
<!------------------------------------------------- -->
<%sub saysth()
name=request.QueryString("name")%>
<hr size="1" width=700>
<form method="post" action="save.asp">
<table border="0" cellpadding="0" cellspacing="0" 
       style="border-collapse: collapse; " 
       width=600 align=center>
<tr><td>
��������:<br>
������<input type="text" name="name" size="23" value=<%=name%>>
���룺<input type="password" name="code" size="28" ID="Password1" style="font-size:12px">
���<a href="register.asp">ע��</a><br>
���⣺<input type="text" name="title" size="67" ID="Text2"><br>
���ݣ�<br>
<textarea rows="11" name="body" cols="72" ID="Textarea1"></textarea>
<br>
<input type="submit" value="�ύ" name="B1" ID="Submit1"><input type="reset" value="����" name="B2" ID="Reset1">
<input type="hidden" name="name_b" size="20" ID="Hidden1"> <input type="hidden" name="title_b" size="28" ID="Hidden2">
</td></tr>
</table>
</form>
<%end sub%>
<html>
<head>
<style type="text/css">
	.top{
	vertical-align:top;}
	.btn{
	background-color:#CCFFCC;
	width:130px;
	height:30px;
	}

</style>
<title>��������ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
        <td id="tbul">��</td>
        <td id= "tbum"><%headtext()%></td>
        <td id="tbur">��</td>
      </tr>
      <tr>
        <td id="tbml">��</td>
        <td id="tbmm"><%maintalk()%><%saysth()%></td>
        <td id="tbmr">��</td>
      </tr>
  
      <tr>
        <td id="tbll">��</td>
        <td id="tblm">��</td>
        <td id="tblr">��</td>
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