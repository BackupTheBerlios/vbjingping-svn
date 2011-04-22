<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->
<%sub maintalk()%>
<%
set showbbs = server.CreateObject("ADODB.recordset")
showbbsStr="select*from bbs order by wtime desc"
showbbs.open showbbsStr,conn,3,2
if showbbs.EOF and showbbs.BOF then%>
<br><center>暂时还没有文章,现在<a href="say.asp">发表</a></center>
<%else
showbbs.PageSize=10'分页
PageN=10 '显示10个页数.
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
      作者:<a href="authordetail.asp?auname=<%=showbbs("name")%>"><%=showbbs("name")%></a>&nbsp;
      <%if showbbs("countwb")=0 then%>
       回复&lt;<%=showbbs("countwb")%>&gt;<!--回复条数 -->
      <%else%>
       <a href="show.asp?id=<%=showbbs("id")%>">
       回复&lt;<font color=#ff0000><b><%=showbbs("countwb")%></b></font>&gt;
       </a><!--回复条数 -->
      <%end if%>&nbsp;
  
      <%wt=showbbs("wtime")'如果为撰写当天的时间，红色显示。
      if year(now())=year(wt) and month(now())=month(wt) and day(now())=day(wt) then%>
       撰写/修改时间:<font color=#ff0000><%=wt%></font>&nbsp;
      <%else%>
       撰写/修改时间:<%=wt%>&nbsp;
      <%end if%>
  
      <a href="modify.asp?id=<%=showbbs("id")%>">修改</a>
      <a href="del.asp?id=<%=showbbs("id")%>">删除</a>
</div>

<hr size="1" width=600  align="left" style=" border:dashed #999999 ">

<%showbbs.movenext
if showbbs.EOF then exit for
Next%>
<center>
<br>
<%
if Page<>1 then%>
<a href="index.asp?Page=1&CurrentPageN=1">(首页)</a>
<%else%>
(当前为首页)
<%end if%>
<%if PageCount>PageN and CurrentPageN>=2 then%>
<a href="index.asp?Page=<%=(CurrentPageN-2)*PageN+1%>&CurrentPageN=<%=CurrentPageN-1%>">(上<%=PageN%>页)</a>
<%end if%>
<%for i=(CurrentPageN-1)*PageN+1 to (CurrentPageN-1)*PageN+PageN%>
<%if i = Page then%>
      <font color=#ff0000 >-<%=i%>-</font>
<%elseif i<showbbs.RecordCount/showbbs.PageSize +1 then%>
      <a href="index.asp?page=<%=i%>&CurrentPageN=<%=CurrentPageN%>">&nbsp;<%=i%>&nbsp;</a>
<%end if%>
<%Next%>
<%if PageCount>PageN and CurrentPageN<PageCount/PageN then%>
<a href="index.asp?Page=<%=CurrentPageN*PageN+1%>&CurrentPageN=<%=CurrentPageN+1%>">(下<%=PageN%>页)</a>
<%end if%>
<%if PageCount>PageN and Page<>PageCount then%>
<a href="index.asp?Page=<%=PageCount%>&CurrentPageN=<%=int(PageCount/PageN)+1%>">(最后一页)</a>
<%end if%>
<%if PageCount=Page then%>(当前为最后页)<%end if%>
(共<%=PageCount%>页)
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
发表文章:<br>
大名：<input type="text" name="name" size="23" value=<%=name%>>
密码：<input type="password" name="code" size="28" ID="Password1" style="font-size:12px">
点此<a href="register.asp">注册</a><br>
标题：<input type="text" name="title" size="67" ID="Text2"><br>
内容：<br>
<textarea rows="11" name="body" cols="72" ID="Textarea1"></textarea>
<br>
<input type="submit" value="提交" name="B1" ID="Submit1"><input type="reset" value="重置" name="B2" ID="Reset1">
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
<title>讨论区主页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (sub01改.psd) -->
<table id="__01" width="950" height="793" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="3">
			<img src="images/sub01改_01.gif" width="814" height="90" alt=""></td>
		<td colspan="2">
			<img src="images/sub01改_02.gif" width="136" height="90" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/sub01_03.gif" width="950" height="68" alt=""></td>
	</tr>
	<tr>
		<td><a href="index.asp"><img src="images/left_menu01.gif" width="219" height="29" alt="" border="0"></a></td>
		<td colspan="4" rowspan="2">
			<img src="images/sub01改_05.gif" width="731" height="49" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/left_menu02.gif" width="219" height="20" alt=""></td>
	</tr>
	<tr>
		<td rowspan="2" class="top">
			<img src="images/sub01改_07.gif" width="219" height="505" alt=""></td>
		<td rowspan="2">
			<img src="images/sub01改_08.gif" width="21" height="505" alt=""></td>
		<td colspan="2" background="images/txt.gif">
		<table width="770" border="0" cellspacing="0" align="center">
      <tr>
        <td id="tbul">　</td>
        <td id= "tbum"><%headtext()%></td>
        <td id="tbur">　</td>
      </tr>
      <tr>
        <td id="tbml">　</td>
        <td id="tbmm"><%maintalk()%><%saysth()%></td>
        <td id="tbmr">　</td>
      </tr>
  
      <tr>
        <td id="tbll">　</td>
        <td id="tblm">　</td>
        <td id="tblr">　</td>
      </tr>
</table>
<!--#include file="footer.htm"--></td>
		<td rowspan="2">
			<img src="images/sub01改_10.gif" width="18" height="505" alt=""></td>
	</tr>
	<tr>
		<td colspan="2">
			<img src="images/sub01改_11.gif" width="692" height="14" alt=""></td>
	</tr>
	<tr>
		<td colspan="5">
			<img src="images/sub01改_12.gif" width="950" height="80" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/分隔符.gif" width="219" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="21" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="574" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="118" height="1" alt=""></td>
		<td>
			<img src="images/分隔符.gif" width="18" height="1" alt=""></td>
	</tr>
</table>
<!-- End ImageReady Slices -->
</body>
</html>