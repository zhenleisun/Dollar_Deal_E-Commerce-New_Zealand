<%

Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'�Զ�����Ҫ���˵��ִ�,�� "��" �ָ�
Fy_In = "'��;��and��exec��insert��select��delete��update��count��*��%��chr��mid��master��truncate��char��declare"
Fy_Inf = split(Fy_In,"��")
If Request.Form<>"" Then
For Each Fy_Post In Request.Form

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('��������վ֣����������\n\n�벻Ҫ�ڲ����а����Ƿ��ַ�����ע�룡');</Script>"
Response.Write "�Ƿ�������ϵͳ�������¼�¼��<br>"
Response.Write "�����ɣУ�"&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "����ʱ�䣺"&Now&"<br>"
Response.Write "����ҳ�棺"&Request.ServerVariables("URL")&"<br>"
Response.Write "�ύ��ʽ���Уϣӣ�<br>"
Response.Write "�ύ������"&Fy_Post&"<br>"
Response.Write "�ύ���ݣ�"&Request.Form(Fy_Post)
Response.End
End If
Next
Next
End If
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('��������վ֣����������\n\n�벻Ҫ�ڲ����а����Ƿ��ַ�����ע�룡');</Script>"
Response.Write "�Ƿ�������JHACKJ�Ѿ������������¼�¼��<br>"
Response.Write "�����ɣУ�"&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "����ʱ�䣺"&Now&"<br>"
Response.Write "����ҳ�棺"&Request.ServerVariables("URL")&"<br>"
Response.Write "�ύ��ʽ���ǣţ�<br>"
Response.Write "�ύ������"&Fy_Get&"<br>"
Response.Write "�ύ���ݣ�"&Request.QueryString(Fy_Get)
Response.End
End If
Next
Next
End If
%>