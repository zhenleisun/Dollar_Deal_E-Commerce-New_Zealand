<!--#include file="Inc/SysConfig.Asp"-->
<%
Dim Content
YxBBs.Head("��̳����")
Select Case Request.QueryString("Action")
Case "ubb"
	ubb
Case "Grade"
	Grade
Case Else
	Main
End Select
YxBBs.Footer()
Set YxBBs=Nothing

Sub Main
	Dim Content
Content="<br><br><b>  === ������� ===</b>"&_
	"<ul><li><b>�ʣ��������ܼ�����̳��</b></li>"&_
"<ul><li>�������Ե����̳������'ע��'ע��Ϊ��վ��Ա,�뽫��������������дЧ,�Ա���������ע���ʹ�������һع���.</li>"&_
        "<li>����Ȼ���Բ���ע��Ϊ���ǵĻ�Ա,����Ϊ�����ܹ�ʹ�ñ���̳��ȫ������,�����Խ�����ע��.</li></ul></ul>"&_
	"<ul><li><b>��������֪���Լ��Ļ��ֺͻ��ֵ����������</b></li>"&_
	"<ul><li>ֻҪ����̳���ҵ��Լ����û�����������ɲ鿴�Լ��Ļ��֡�</li>"&_
	"<li>������̳������������ͨ������̳�˵���<a href=Members.Asp?Action=top><font color=bule>�û��б�</font></a>�鿴��</li></ul></ul>"&_
	"<ul><li><b>�������������Ӻ����ǩ���� </b></li>"&_
	"<ul><li>������ͨ����̳������YourName���¡��޸����ϡ�����ǩ��һ����д�����ĸ���ǩ����</li>"&_
	"<li>ǩ��֧��UBB������ʹ��ͼƬ����ʽ��<font color=bule>[img]ͼƬ��ַ[/img]</font>��</li></ul></ul>"&_
	"<ul><li><b>��ο����ҵ���Ҫ�����£�</b></li>"&_
	"<ul><li>������ʹ����̳���������ܣ�����������̳�������ڡ���桱�µġ�<a href=Search.Asp><font color=bule>��̳����</font></a>����д���������������������</li>"&_
        "<li>�����ڸ������·�Ҳ���԰�����������������������ѡ��ʹ�á������������������������������������ظ������Ƚ��м���������</li></ul></ul>"&_
	"<ul><li><b>ʲô�Ǿ�������������˭�������뾫�����ģ���β鿴��</b>"&_
	"<ul><li>����������̳����������м�ֵ�����������ϸ߻����ݱȽ�����������ӵģ�����ͨ�������ھ��������ҵ��ܶ����õĶ���������̳��ÿ�����涼���Լ��ľ��������������ɰ����������������Խ������ϵ����Ӽ��뵽�����������ɱ༭�����ټӹ���</li>"&_
	"<li>ֻҪ������ذ��棬�ڰ�������Ϸ��Ϳ��Կ��������澫����������������ɲ鿴���ڡ���桱��Ҳ��һ������������������ȫ����̳�ľ������ڣ����������Դ��ַ�ʽ�鿴�������������ؾ����뵽��ذ�����ҡ�</li></ul></ul>"&_
	"<ul><li><b>��Ϊ������������ʲô������ЩȨ����</b></li>"&_
	"<ul><li>������Ȼ��������֮��������Ϊ�����ģ�Ը��Ϊ�����޳�����</li>"&_
	"<li>ӵ��һ��Ⱥ�ڻ������ܻ�ʱ��ά����̳�ġ�</li>"&_
	"<li>�������Բ�ѯ��������ӣ���������ɾ����༭��������ӡ�</li>"&_
	"<li>�������԰����Ӽ��������������ö�������ö����������ͳ������ӣ����뾫���������������������������</li>"&_ 
	"<li>�������������������Ϸ�����ɾ�����档</li>"&_
	"<li>�����̳�����˰����̳й��ܣ��ϼ����������Թ����¼���̳�����ӡ�</li></ul></ul>"&_
        "<ul><li><b>������������ô�죿</b></li>"&_
       "<ul><li>����������룬�������ע��ʱ��д������������������<font color=bule>ȡ������</font></a>��</ul><ul><li>�����������������ʾ������𰸣�������̳����Ա��ϵ��������Ϊ���趨�µ����롣</li></ul><br>"&_
"<br><li><font face=Wingdings color=#9ec750>v</font><a href=?Action=ubb><font color=red><b>��ذ�����UBB��ǩ����</b></font></a></li><font face=Wingdings color=#9ec750>v</font><br><li><font face=Wingdings color=#9ec750>v</font><a href=?Action=Grade><font color=red><b>��ذ�������̳�ȼ�����</b></font></a><font face=Wingdings color=#9ec750>v</font></li></ul></blockquote>"
Call YxBBs.ShowTable("�������",Content)
End sub

sub ubb
Content=Content&"<br><br><b>=== UBB��ǩ���� ===</b>"&_
	"<ul><li><b>UBB��ǩ���ǲ�����ʹ��HTML�﷨������£�ͨ����̳������ת��������������֧���������õġ���Σ���Ե�HTMLЧ����ʾ������Ϊ����ʹ��˵����</b></li>"&_
	"<br><br><ul><li><p><font color=red>[B]</font><b>����</b><font color=red>[/B]</font><br>�����ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ����Ч����</li>"&_
	"<p><li><font color=red>[I]</font><i>����</i><font color=red>[/I]</font><br>�����ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪб��Ч����</li>"&_
	"<p><li><font color=red>[U]</font><u>����</u><font color=red>[/U]</font><br>�����ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ�»���Ч����</li>"&_
	"<P><li><font color=red>[URL=http://www.yimxu.com]</font><A HREF=http://www.yimxu.com>Y��</A><font color=red>[/URL]</font>��<br>������Ҳ�������Ӿ����ַ�����������ӡ�</li>"&_
	"<P><li><font color=red>[img]</font>http://www.yimxu.com/logo.gif<font color=red>[/img]</font><br>�ڱ�ǩ���м����ͼƬ��ַ����ʵ�ֲ�ͼЧ����</li>"&_
	"<P><li><font color=red>[fly]</font>����<font color=red>[/fly]</font><br>�ڱ�ǩ���м�������ֿ���ʵ�����ַ���Ч�������������ơ�</li>"&_
	"<P><li><font color=red>[color=��ɫ����]</font>����<font color=red>[/color]</font><br>����������ɫ���룬�ڱ�ǩ���м�������ֿ���ʵ��������ɫ�ı䡣</li><br><br>"&_
        "<li><font face=Wingdings color=#9ec750>v</font><a href=?Action=main><font color=red><b>��ذ�������̳��������</b></font></a></li><font face=Wingdings color=#9ec750>v</font><br><li><font face=Wingdings color=#9ec750>v</font><a href=?Action=Grade><font color=red><b>��ذ�������̳�ȼ�����</b></font></a><font face=Wingdings color=#9ec750>v</font></li><br><br>"
	 Call YxBBs.ShowTable("UBB��ǩ",Content)
End sub
sub Grade
Dim Rs
	Content="<tr><td height=25>��Ա�ȼ�</td><td>�ȼ��ƺ�</td><td>��������</td><td>�ȼ�ͼƬ</td></tr>"
	Set Rs=YxBBs.Execute("Select GradeNum,GradeName,GradePic,UserMinPostNum from YX_UserGrade Order By GradeNum")
	Do while Not Rs.eof
		Content=Content&"<tr><td height=25>"&Rs(0)&"</td><td>"&Rs(1)&"</td><td>"&Rs(3)&"</td><td><img src="""&Rs(2)&"""></td></tr>"
		Rs.MoveNext
	Loop
	Content=Content&"<tr><td colspan=4>&nbsp;&nbsp;�ر�˵����<br>&nbsp;&nbsp;<b><font color=red>18</font></b>��Ϊ��̳����<br>&nbsp;&nbsp;<b><font color=red>19</font></b>��Ϊǰ̨����Ա<br>&nbsp;&nbsp;<b><font color=red>20</font></b>��Ϊ��̳վ��<br> </font><br><br><li><font face=Wingdings color=#9ec750>v</font><a href=?Action=main><b>��ذ�������̳��������</b></a></li><font face=Wingdings color=#9ec750>v</font><br><li><font face=Wingdings color=#9ec750>v</font><a href=?Action=ubb><b>��ذ�����UBB��ǩ����</b></a><font face=Wingdings color=#9ec750>v</font></li><br><br>"
    Call YxBBs.ShowTable("��̳�ȼ�",Content)
End sub
%>