<html>
<head>
</head>
<script language = javascript>

function ok() {
  if (frm_up.text_title.value.length == 0) {
   alert("������ �Է����ּ���.")
   form_up.txt_title.focus();
      }

  else if (frm_up.text_name.value.length == 0) {
   alert("�̸��� �����ּ���.")
   form_up.txt_name.focus();
      }

  else if (frm_up.text_area.value.length == 0) {
   alert("������ �����ּ���.")
   frm_up.text_area.focus();
      }

  else if(frm_up.text_pw.value.length <4) {
   alert("������ ���� PW 4�ڸ� �Է����ּ���")
   frm_up.text_pw.focus();
  }

  else {
    frm_up.submit();
      }
    }

</script>
<body onload="frm_up.text_title.focus();">
<form action="board_input.asp" method="get" name="frm_up">
  <center>
  <table border="0" width="80%">
  <tr>
    <td colspan="2">����<input type="text" name="text_title" maxlength="40" style="width: 100%;">
    </td>
  </tr>

  <tr>
    <td>�ۼ��� <br><input type="text" name="text_name" maxlength="6" readonly
    value=<% if Session("user_id") = "" then
    response.write "��ȸ��"
    else
    response.write session("user_id")
    end if %>>
    </td>

   <tr>
     <td>��й�ȣ <br><input type="text" name="text_pw" maxlength="4" placeholder="��й�ȣ 4�ڸ� �Է�">
     </td>
   </tr>

  </tr>
	<td>����</td>
  </table>
  <textarea name="text_area" rows=15% cols=70%></textarea><p>
  <input type="button" value="�ڷ�" onclick="history.go(-1);">
  <input type="button" value="�Է�" onclick="ok();">
  </center>
  </form>
  </body>
</html>
