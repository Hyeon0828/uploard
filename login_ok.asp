<html>
<script language=javascript>

function logout() {
window.open('logout.asp', '_self');
}

function info() {
window.open('info.asp', 'new', 'width=400, height=350, status=no, scrollbars=auto');
}
</script>

<%
if Session("user_id") = "" then
 response.write "<script>alert('���ѻ����� �����ϴ�.'); history.go(-1);</script>"
end if
%>

<body>
<form name="login_ok" method="post">
<table>
<div id="name">
 <% response.write Session("user_id") %>�� ȯ���մϴ�.<p><br>
 </div><p>
  <tr>
    <td><font size="2">���ӽð� : <%=now()%></font></td>
  </tr>
<div>
 <input type="button" value="My Info" onclick="info();">
 <input type="button" value="Log-out" onclick="logout();">
</div>
</table>


 </form>
 </html>
