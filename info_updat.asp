<html>
<%
Db_id = Session("user_id")
Db_name = request("text_name")
Db_zip1 = request("zip1")
Db_zip2 = request("zip2")
Db_zip3 = request("zip3")
Db_zip4 = request("zip4")
Db_email1 = request("email")
Db_email2 = request("domain")
Db_phone1 = request("phone1")
Db_phone2 = request("phone2")
Db_phone3 = request("phone3")

Set db = Server.createObject ("ADODB.Connection")
  dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
  db.Open(dbMain)

update_db = "update ex1table set db_name = '" & Db_name & "', db_phone1 = '" & Db_phone1 & "', db_phone2 = '" & Db_phone2 & "', db_phone3 = '" & Db_phone3 & "', db_email1 = '" & Db_email1 & "', db_email2 = '" & Db_email2 & "', db_zip1 = '" & Db_zip1 & "', db_zip2 = '" & Db_zip2 & "', db_zip3 = '" & Db_zip3 & "', db_zip4 = '" & Db_zip4 & "' where db_id = '" & Db_id & "'"
db.execute(update_db)
db.close()
%>

<body>
</body>
<script language="javascript">
alert("회원정보 수정완료")
window.Open("info_asp", "_self")
self.close();
</script>
</html>
