<%
Dim strID

strID = request("IDtext1")

Set db = Server.createObject ("ADODB.Connection")
 con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
 db.Open(con)

select_sql = "select db_id from ex1table where db_id = '" & strID & "'"
set rs=db.execute(select_sql)
%>

<head>
<title>ID 중복확인</title>
<script language="javascript">

function id_submit(arg) {
 if(arg != "") {
 id_check.submit();
 }
 else {
 alert("ID를 먼저 입력하세요.")
 id_check.IDtext1.focus();
 return false;
 }
}

function set_id() {
opener.document.frm.IDtext1.value;
opener.frm.IDtext1.readOnly = true;
self.close();
}

</script>
</head>


<body onload="this.document.id.check.IDtext1.focus();">
<table width="300" border="0" cellspacing="0" cellpadding="0">
<tr>
<td>
<table width="95%" border="1" cellspacing="0" cellpadding="5" align="center" height="150">

 <tr>
 <td colspan="2"  height="40">
 <div align="center">
 <font size="5">
 <b>ID 중복확인</b>
 </font>
 </div>
 </td>
 </tr>

<form method="post" name="id_check" action="id_check.asp" onsubmit="return id_submit(id_check.IDtext1.value)">
<tr>
<td colspan="2" height="40">
 <div align="center">
 <input type="text" name="IDtext1" size="10" maxlength="9" value="<%=strID%>">
 <input type="submit" name="submit" value="검색" onclick="id_submit(id_check.IDtext1.value)">
 </div>
</td>
</tr>
</form>

<tr>
<td colspan="2" align="center">

<%
 IF NOT rs.EOF OR NOT rs.BOF THEN
%>

<font size="2">입력하신 ID [<font face="굴림" color="#FF0000"><%=strID%></font>]는 <br><br>이미 사용중인 ID입니다.<br>
<%
else
%>
<font size="2">입력하신 ID [<font face="굴림" color="#FF0000"><%=strID%></font>]는 <br><br> 사용 가능한 ID입니다.
<p>
<input type="button" name="submit2" value="ID 사용" onclick="set_id()"><br>
</p>
<%
end if
 rs.close
 set rs = nothing
 %>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>
</html>
