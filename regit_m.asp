<html>
  <head>
     <meta http-equiv="X-UA-Compatible" content="IE=edge">
     <meta name="viewport" content="width=device-width, initial-scale=1">
  </head>

 <style type="text/css">
    #div_tyme{
      float: right;
      display: flex;
    }
    #div_main{
      width: 660px;
      height: 400px;
      margin-top: 20px;
    }
   .box_title{
     border: 0px;
     background-color: white;
   }
   .box_font{
     font-size: 17px;
     margin-right: 10px;
   }
   .table {
     width: 100%;
     height: 450px;
   }
 </style>

 <body>
   <center>
    <form name="form_a2">

      <div id="div_main">
        <h1><b>�ִ� �˻�</b></h1><br>

      <table class="table">
        <thead>
          <tr>
            <th>���̵�</th>
            <th>����</th>
            <th>����</th>
            <th>�������</th>
            <th>�˻�Ƚ��</th>
          </tr>
        </thead>
        <tbody align=center>
       <%
         Dim cnt_id, sql
         Set db = Server.createObject ("ADODB.Connection")
         dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
         db.Open (dbMain)

         sql = "select top 5 * from ex1table ORDER BY db_m DESC"
         Set search_db = db.execute(sql)

          Do while Not search_db.EOF

            db_id = search_db("db_id")
            db_name = search_db("db_name")
            db_gender = search_db("db_gender")
            db_birth = search_db("db_birth")
            db_m = search_db("db_m")

             response.write "<tr><td>" & db_id & "</td><td>" & db_name & "</td><td>" & db_gender & "</td><td>" & db_birth & "</td><td>" & db_m & "</td></tr>"

             search_db.MoveNext
           Loop

         db.close()
       %>
     </table>
  </form>
  </body>
 </html>
