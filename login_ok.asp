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
 response.write "<script>alert('권한사항이 없습니다.'); history.go(-1);</script>"
end if
%>

<body>
<form name="login_ok" method="post">
<table>
<div id="name">
 <% response.write Session("user_id") %>님 환영합니다.<p><br>
 </div><p>
  <tr>
    <td><font size="2">접속시간 : <%=now()%></font></td>
  </tr>
<div>
 <input type="button" value="My Info" onclick="info();">
 <input type="button" value="Log-out" onclick="logout();">
</div>
</table>


 </form>
 </html>
