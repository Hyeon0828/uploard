<html>
<%
 db_title = request("text_title")
 db_name = request("text_name")
 db_main = request("text_main")
 db_date = now()

   Server.createObject ("ADODB.Connection")
   dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("board.accdb")
   db.Open(dbMain)

   update_db = "update board_table set db_title = '" & db_title & "', db_name = '" & db_name & "', db_main = '" & db_main & "', db_date = '" & db_date & "'  where db_title = '" & db_title & "'"
   db.Execute(update_db)
   db.close()
 %>

  <body>
  </body>
  <script language = "javascript">
   alert("수정완료.")
   history.go(-1);
  </script>
</html>
