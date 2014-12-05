var IMAGE_PATH;
var ICON_PATH;
var IMAGE_ATTACH_PATH;
var IMAGE_UPLOAD_CGI;
var MENU_BORDER_COLOR;
var MENU_BG_COLOR;
var MENU_TEXT_COLOR;
var MENU_SELECTED_COLOR;
var TOOLBAR_BG_COLOR;
var FORM_BORDER_COLOR;
var FORM_BG_COLOR;
var OBJ_NAME;
var SELECTION;
var RANGE;
var RANGE_TEXT;
var EDITFORM_DOCUMENT;
var UPLOAD_DOCUMENT;
var BROWSER;
var TOOLBAR_ICON;

var MSG_INPUT_URL = "请输入正确的URL地址。";
var MSG_INVALID_IMAGE = "只能选择GIF,JPG,PNG,BMP格式的图片，请重新选择。";
var MSG_INVALID_MEDIA = "只能选择MP3,WMV,ASF,WMA,RM格式的多媒体，请重新选择。";
var MSG_SELECT_TARGET = "请选择目标文本。";
var MSG_INVALID_WIDTH = "宽度不是数字，请重新输入。";
var MSG_INVALID_HEIGHT = "高度不是数字，请重新输入。";
var MSG_INVALID_BORDER = "边框不是数字，请重新输入。";
var STR_WIDTH = "宽";
var STR_HEIGHT = "高";
var STR_BORDER = "边";
var STR_BUTTON_CONFIRM = "确定";
var STR_BUTTON_CANCEL = "取消";
var STR_BUTTON_PREVIEW = "预览";
var STR_BUTTON_LISTENING = "试听";
var STR_IMAGE_LOCAL = "本地";
var STR_IMAGE_REMOTE = "远程";
var STR_LINK_BLANK = "新窗口";
var STR_LINK_NOBLANK = "当前窗口";
var STR_LINK_TARGET = "目标";

var EDITOR_FONT_FAMILY = "SimSun";

var FONT_NAME = Array(
					Array('SimSun', '宋体'), 
					Array('SimHei', '黑体'), 
					Array('FangSong_GB2312', '仿宋体'), 
					Array('KaiTi_GB2312', '楷体'), 
					Array('NSimSun', '新宋体'), 
					Array('Arial', 'Arial'), 
					Array('Arial Black', 'Arial Black'), 
					Array('Times New Roman', 'Times New Roman'), 
					Array('Courier New', 'Courier New'), 
					Array('Tahoma', 'Tahoma'), 
					Array('Verdana', 'Verdana'), 
					Array('GulimChe', 'GulimChe'), 
					Array('MS Gothic', 'MS Gothic') 
					);
var ZOOM_TABLE = Array('250%', '200%', '150%', '120%', '100%', '80%', '50%');
var TITLE_TABLE = Array(
					Array('H1', '标题 1'), 
					Array('H2', '标题 2'), 
					Array('H3', '标题 3'), 
					Array('H4', '标题 4'), 
					Array('H5', '标题 5'), 
					Array('H6', '标题 6')
					);
var FONT_SIZE = Array(
					Array(1,'8pt'), 
					Array(2,'10pt'), 
					Array(3,'12pt'), 
					Array(4,'14pt'), 
					Array(5,'18pt'), 
					Array(6,'24pt'), 
					Array(7,'36pt')
				);
var SPECIAL_CHARACTER = Array('§','№','☆','★','○','●','◎','◇','◆','□','℃','‰','■','△','▲','※',
							'→','←','↑','↓','〓','¤','°','＃','＆','＠','＼','︿','＿','￣','―','α',
							'β','γ','δ','ε','ζ','η','θ','ι','κ','λ','μ','ν','ξ','ο','π','ρ',
							'σ','τ','υ','φ','χ','ψ','ω','≈','≡','≠','＝','≤','≥','＜','＞','≮',
							'≯','∷','±','＋','－','×','÷','／','∫','∮','∝','∞','∧','∨','∑','∏',
							'∪','∩','∈','∵','∴','⊥','∥','∠','⌒','⊙','≌','∽','〖','〗','【','】');
var TOP_TOOLBAR_ICON = Array(
						Array('Html_SOURCE', 'source.gif', '42', '20', '视图转换'),
						Array('Html_ZOOM', 'zoom.gif', '20', '20', '显示比例'),
						Array('Html_COPY', 'copy.gif', '20', '20', '复制'),
						Array('Html_PASTE', 'paste.gif', '20', '20', '粘贴'),
						Array('Html_SELECTALL', 'selectall.gif', '20', '20', '全选'),
						Array('Html_JUSTIFYLEFT', 'justifyleft.gif', '20', '20', '左对齐'),
						Array('Html_JUSTIFYCENTER', 'justifycenter.gif', '20', '20', '居中'),
						Array('Html_JUSTIFYRIGHT', 'justifyright.gif', '20', '20', '右对齐'),
						Array('Html_JUSTIFYFULL', 'justifyfull.gif', '20', '20', '两端对齐'),
						Array('Html_NUMBEREDLIST', 'numberedlist.gif', '20', '20', '编号'),
						Array('Html_UNORDERLIST', 'unorderedlist.gif', '20', '20', '项目符号'),
						Array('Html_INDENT', 'indent.gif', '20', '20', '减少缩进'),
						Array('Html_OUTDENT', 'outdent.gif', '20', '20', '增加缩进'),
						Array('Html_SUBSCRIPT', 'subscript.gif', '20', '20', '下标'),
						Array('Html_SUPERSCRIPT', 'superscript.gif', '20', '20', '上标'),
                                                Array('Html_SPECIALCHAR', 'specialchar.gif', '20', '20', '特殊字符'),
						Array('Html_DATE', 'date.gif', '20', '20', '日期'),
						Array('Html_TIME', 'time.gif', '20', '20', '时间')
				  );
var BOTTOM_TOOLBAR_ICON = Array(
						Array('Html_TITLE', 'title.gif', '28', '20', '标题'),
						Array('Html_FONTNAME', 'font.gif', '28', '20', '字体'),
						Array('Html_FONTSIZE', 'fontsize.gif', '28', '20', '文字大小'),
						Array('Html_TEXTCOLOR', 'textcolor.gif', '20', '20', '文字颜色'),
						Array('Html_BGCOLOR', 'bgcolor.gif', '20', '20', '文字背景'),
						Array('Html_BOLD', 'bold.gif', '20', '20', '粗体'),
						Array('Html_ITALIC', 'italic.gif', '20', '20', '斜体'),
						Array('Html_UNDERLINE', 'underline.gif', '20', '20', '下划线'),
						Array('Html_STRIKE', 'strikethrough.gif', '20', '20', '删除线'),
						Array('Html_REMOVE', 'removeformat.gif', '20', '20', '删除格式'),
						Array('Html_IMAGE', 'image.gif', '20', '20', '图片'),
						Array('Html_FLASH', 'flash.gif', '20', '20', 'Flash'),
						Array('Html_MEDIA', 'media.gif', '20', '20', '视频'),
						Array('Html_LAYER', 'layer.gif', '20', '20', '层'),
						Array('Html_TABLE', 'table.gif', '20', '20', '表格'),
						
						Array('Html_HR', 'hr.gif', '20', '20', '横线'),
						Array('Html_ICON', 'emoticons.gif', '20', '20', '笑脸'),
						Array('Html_LINK', 'link.gif', '20', '20', '创建超级连接'),
						Array('Html_UNLINK', 'unlink.gif', '20', '20', '删除超级连接')

				  );
var COLOR_TABLE = Array(
						"#FF0000", "#FFFF00", "#00FF00", "#00FFFF", "#0000FF", "#FF00FF", "#FFFFFF", "#F5F5F5", "#DCDCDC", "#FFFAFA",
						"#D3D3D3", "#C0C0C0", "#A9A9A9", "#808080", "#696969", "#000000", "#2F4F4F", "#708090", "#778899", "#4682B4",
						"#4169E1", "#6495ED", "#B0C4DE", "#7B68EE", "#6A5ACD", "#483D8B", "#191970", "#000080", "#00008B", "#0000CD",
						"#1E90FF", "#00BFFF", "#87CEFA", "#87CEEB", "#ADD8E6", "#B0E0E6", "#F0FFFF", "#E0FFFF", "#AFEEEE", "#00CED1",
						"#5F9EA0", "#48D1CC", "#00FFFF", "#40E0D0", "#20B2AA", "#008B8B", "#008080", "#7FFFD4", "#66CDAA", "#8FBC8F",
						"#3CB371", "#2E8B57", "#006400", "#008000", "#228B22", "#32CD32", "#00FF00", "#7FFF00", "#7CFC00", "#ADFF2F",
						"#98FB98", "#90EE90", "#00FF7F", "#00FA9A", "#556B2F", "#6B8E23", "#808000", "#BDB76B", "#B8860B", "#DAA520",
						"#FFD700", "#F0E68C", "#EEE8AA", "#FFEBCD", "#FFE4B5", "#F5DEB3", "#FFDEAD", "#DEB887", "#D2B48C", "#BC8F8F",
						"#A0522D", "#8B4513", "#D2691E", "#CD853F", "#F4A460", "#8B0000", "#800000", "#A52A2A", "#B22222", "#CD5C5C",
						"#F08080", "#FA8072", "#E9967A", "#FFA07A", "#FF7F50", "#FF6347", "#FF8C00", "#FFA500", "#FF4500", "#DC143C",
						"#FF0000", "#FF1493", "#FF00FF", "#FF69B4", "#FFB6C1", "#FFC0CB", "#DB7093", "#C71585", "#800080", "#8B008B",
						"#9370DB", "#8A2BE2", "#4B0082", "#9400D3", "#9932CC", "#BA55D3", "#DA70D6", "#EE82EE", "#DDA0DD", "#D8BFD8",
						"#E6E6FA", "#F8F8FF", "#F0F8FF", "#F5FFFA", "#F0FFF0", "#FAFAD2", "#FFFACD", "#FFF8DC", "#FFFFE0", "#FFFFF0",
						"#FFFAF0", "#FAF0E6", "#FDF5E6", "#FAEBD7", "#FFE4C4", "#FFDAB9", "#FFEFD5", "#FFF5EE", "#FFF0F5", "#FFE4E1"
					);
function cleanHtml(content)
{
	content = content.replace(/<p>&nbsp;<\/p>/gi,"")
	content = content.replace(/<p><\/p>/gi,"<p>")
	content = content.replace(/<div><\/\1>/gi,"")
	content = content.replace(/<p>/,"<br>")
	content = content.replace(/(<(meta|iframe|frame|span|tbody|layer)[^>]*>|<\/(iframe|frame|meta|span|tbody|layer)>)/gi, "");
	content = content.replace(/<\\?\?xml[^>]*>/gi, "") ;
	content = content.replace(/o:/gi, "");
	content = content.replace(/&nbsp;/gi, " ");
return content;
}
//代码过滤及JS提取
function FilterScript(content)
{
	content = content.replace(/<(\w[^div|>]*) class\s*=\s*([^>|\s]*)([^>]*)/gi,"<$1$3") ;
	content = content.replace(/<(\w[^font|>]*) style\s*=\s*\"[^\"]*\"([^>]*>)/gi,"<$1 $2") ;
	content = content.replace(/<(\w[^>]*) lang\s*=\s*([^>|\s]*)([^>]*)/gi,"<$1$3") ;
	var RegExp = /<(script[^>]*)>(.*)<\/script>/gi;
	content = content.replace(RegExp, "[code]&lt;$1&gt;$2&lt;script&gt;[\/code]");
	RegExp = /<(\w[^>|\s]*)([^>]*)(on(finish|mouse|Exit|error|click|key|load|change|focus|blur))(.[^>]*)/gi;
	content = content.replace(RegExp, "<$1")
	RegExp = /<(\w[^>|\s]*)([^>]*)(&#|window\.|javascript:|js:|about:|file:|Document\.|vbs:|cookie| name| id)(.[^>]*)/gi;
	content = content.replace(RegExp, "<$1")
	return content;
}

function HtmlGetBrowser()
{
	var browser = '';
	var agentInfo = navigator.userAgent.toLowerCase();
	if (agentInfo.indexOf("msie") > -1) {
		var re = new RegExp("msie\\s?([\\d\\.]+)","ig");
		var arr = re.exec(agentInfo);
		if (parseInt(RegExp.$1) >= 5.5) {
			browser = 'IE';
		}
	} else if (agentInfo.indexOf("firefox") > -1) {
		browser = 'FF';
	} else if (agentInfo.indexOf("netscape") > -1) {
		var temp1 = agentInfo.split(' ');
		var temp2 = temp1[temp1.length-1].split('/');
		if (parseInt(temp2[1]) >= 7) {
			browser = 'NS';
		}
	} else if (agentInfo.indexOf("gecko") > -1) {
		browser = 'ML';
	} else if (agentInfo.indexOf("opera") > -1) {
		var temp1 = agentInfo.split(' ');
		var temp2 = temp1[0].split('/');
		if (parseInt(temp2[1]) >= 9) {
			browser = 'OPERA';
		}
	}
	return browser;
}
function HtmlGetFileName(file, separator)
{
	var temp = file.split(separator);
	var len = temp.length;
	var fileName = temp[len-1];
	return fileName;
}
function HtmlGetFileExt(fileName)
{
	var temp = fileName.split(".");
	var len = temp.length;
	var fileExt = temp[len-1].toLowerCase();
	return fileExt;
}
function HtmlCheckImageFileType(file, separator)
{
	if (separator == "/" && file.match(/http:\/\/.{3,}/) == null) {
		alert(MSG_INPUT_URL);
		return false;
	}
	var fileName = HtmlGetFileName(file, separator);
	var fileExt = HtmlGetFileExt(fileName);
	if (fileExt != 'gif' && fileExt != 'jpg' && fileExt != 'png' && fileExt != 'bmp') {
		alert(MSG_INVALID_IMAGE);
		return false;
	}
	return true;
}
function HtmlCheckFlashFileType(file, separator)
{
	if (file.match(/http:\/\/.{3,}/) == null) {
		alert(MSG_INPUT_URL);
		return false;
	}
	var fileName = HtmlGetFileName(file, "/");
	var fileExt = HtmlGetFileExt(fileName);
	
	return true;
}
function HtmlCheckMediaFileType(file, separator)
{
	if (file.match(/http:\/\/.{3,}/) == null) {
		alert(MSG_INPUT_URL);
		return false;
	}
	var fileName = HtmlGetFileName(file, "/");
	var fileExt = HtmlGetFileExt(fileName);
	if (fileExt != 'mp3' && fileExt != 'wmv' && fileExt != 'asf' && fileExt != 'wma' && fileExt != 'rm') {
		alert(MSG_INVALID_MEDIA);
		return false;
	}
	return true;
}
function HtmlUpload(url)
{
	UPLOAD_DOCUMENT.uploadForm.action = url;
	UPLOAD_DOCUMENT.uploadForm.target = '_self';
	UPLOAD_DOCUMENT.uploadForm.submit();
}
function HtmlHtmlentities(str)
{
	str = str.replace(/</g,'&lt;');
	str = str.replace(/>/g,'&gt;');
	str = str.replace(/&/g,'&amp;');
	str = str.replace(/"/g,'&quot;');
	return str;
}
function HtmlHtmlentitiesDecode(str)
{
	str = str.replace(/&lt;/g,'<');
	str = str.replace(/&gt;/g,'>');
	str = str.replace(/&amp;/g,'&');
	str = str.replace(/&quot;/g,'"');
	return str;
}
function HtmlHtmlToXhtml(str) 
{
	str = str.replace(/<p([^>]*>)/gi, "<div$1");
	str = str.replace(/<\/p>/gi, "</div>");
	str = str.replace(/<br>/gi, "<br />");
	str = str.replace(/(<\w+)([^>]*>)/gi, function ($0,$1,$2) {
						return($1.toLowerCase() + $2);
					}
				);
	str = str.replace(/(<\/\w+>)/gi, function ($0,$1) {
						return($1.toLowerCase());
					}
				);
	str = HtmlTrim(str);
	return str;
}
function HtmlTrim(str)
{
	str = str.replace(/^\s+|\s+$/g, "");
	str = str.replace(/[\r\n]+/g, "\r\n");
	return str;
}
function HtmlGetTop(id)
{
	var top = 0;
	var tmp = '';
	var obj = document.getElementById(id);
	while (eval("obj" + tmp).tagName != "BODY") {
		tmp += ".offsetParent";
		top += eval("obj" + tmp).offsetTop;
	}
	return top;
}
function HtmlGetLeft(id)
{
	var left = 0;
	var tmp = '';
	var obj = document.getElementById(id);
	while (eval("obj" + tmp).tagName != "BODY") {
		tmp += ".offsetParent";
		left += eval("obj" + tmp).offsetLeft;
	}
	return left;
}
function HtmlClearTemp()
{
	document.getElementById('popupData').innerHTML = '';
	document.getElementById('popupName').innerHTML = '';
}
function HtmlGetMenuCommonStyle(top, left)
{
	var str = 'position:absolute;top:'+top+'px;left:'+left+'px;font-size:12px;color:'+MENU_TEXT_COLOR+
			';background-color:'+MENU_BG_COLOR+';border:solid 1px '+MENU_BORDER_COLOR+';z-index:1';
	return str;
}
function HtmlDrawMenu(mode, content)
{
	var top = HtmlGetTop(mode) + 32;
	var left = HtmlGetLeft(mode) + 1;
	var str = '';
	str += '<div style="'+HtmlGetMenuCommonStyle(top, left)+'">';
	str += content;
	str += '</div>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = mode;
}
function HtmlDrawZoom()
{
	var str = '';
	for (i = 0; i < ZOOM_TABLE.length; i++) {
		str += '<div style="padding:2px;width:120px;cursor:pointer;" ' + 
		'onclick="javascript:HtmlExecute(\'Html_ZOOM_END\', \'' + ZOOM_TABLE[i] + '\');" ' + 
		'onmouseover="javascript:this.style.backgroundColor=\''+MENU_SELECTED_COLOR+'\';" ' +
		'onmouseout="javascript:this.style.backgroundColor=\''+MENU_BG_COLOR+'\';">' + 
		ZOOM_TABLE[i] + '</div>';
	}
	HtmlDrawMenu('Html_ZOOM', str);
}
function HtmlDrawTitle()
{
	var str = '';
	for (i = 0; i < TITLE_TABLE.length; i++) {
		str += '<div style="width:140px;cursor:pointer;" ' + 
		'onclick="javascript:HtmlExecute(\'Html_TITLE_END\', \'' + TITLE_TABLE[i][0] + '\');" ' + 
		'onmouseover="javascript:this.style.backgroundColor=\''+MENU_SELECTED_COLOR+'\';" ' +
		'onmouseout="javascript:this.style.backgroundColor=\''+MENU_BG_COLOR+'\';"><' + TITLE_TABLE[i][0] + ' style="margin:2px;">' + 
		TITLE_TABLE[i][1] + '</' + TITLE_TABLE[i][0] + '></div>';
	}
	HtmlDrawMenu('Html_TITLE', str);
}
function HtmlDrawFontname()
{
	var str = '';
	for (i = 0; i < FONT_NAME.length; i++) {
		str += '<div style="font-family:' + FONT_NAME[i][0] + 
		';padding:2px;width:160px;cursor:pointer;" ' + 
		'onclick="javascript:HtmlExecute(\'Html_FONTNAME_END\', \'' + FONT_NAME[i][0] + '\');" ' + 
		'onmouseover="javascript:this.style.backgroundColor=\''+MENU_SELECTED_COLOR+'\';" ' +
		'onmouseout="javascript:this.style.backgroundColor=\''+MENU_BG_COLOR+'\';">' + 
		FONT_NAME[i][1] + '</div>';
	}
	HtmlDrawMenu('Html_FONTNAME', str);
}
function HtmlDrawFontsize()
{
	var str = '';
	for (i = 0; i < FONT_SIZE.length; i++) {
		str += '<div style="font-size:' + FONT_SIZE[i][1] + 
		';padding:2px;width:120px;cursor:pointer;" ' + 
		'onclick="javascript:HtmlExecute(\'Html_FONTSIZE_END\', \'' + FONT_SIZE[i][0] + '\');" ' + 
		'onmouseover="javascript:this.style.backgroundColor=\''+MENU_SELECTED_COLOR+'\';" ' +
		'onmouseout="javascript:this.style.backgroundColor=\''+MENU_BG_COLOR+'\';">' + 
		FONT_SIZE[i][1] + '</div>';
	}
	HtmlDrawMenu('Html_FONTSIZE', str);
}
function HtmlDrawColorTable(cmdName)
{
	var top = HtmlGetTop(cmdName) + 32;
	var left = HtmlGetLeft(cmdName) + 1;
	var str = '';
	str += '<table cellpadding="0" cellspacing="2" border="0" style="'+HtmlGetMenuCommonStyle(top, left)+'">';
	for (i = 0; i < COLOR_TABLE.length; i++) {
		if (i == 0 || (i >= 10 && i%10 == 0)) {
			str += '<tr>';
		}
		str += '<td style="width:12px;height:12px;border:1px solid #AAAAAA;font-size:10px;cursor:pointer;background-color:' +
		COLOR_TABLE[i] + ';" onmouseover="javascript:this.style.borderColor=\'#000000\';" ' +
		'onmouseout="javascript:this.style.borderColor=\'#AAAAAA\';" ' + 
		'onclick="javascript:HtmlExecute(\''+cmdName+'_END\', \'' + COLOR_TABLE[i] + '\');">&nbsp;</td>';
		if (i >= 9 && i%(i-1) == 0) {
			str += '</tr>';
		}
	}
	str += '</table>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = cmdName;
}
function HtmlCheckSelected()
{
	if (RANGE_TEXT == '') {
		alert(MSG_SELECT_TARGET);
		return false;
	}
}
function HtmlDrawLink()
{
	var top = HtmlGetTop('Html_LINK') + 32;
	var left = HtmlGetLeft('Html_LINK') - 220;
	var str = '';
	str += '<table cellpadding="0" cellspacing="0" style="width:250px;'+HtmlGetMenuCommonStyle(top, left)+'">' + 
		'<tr><td style="width:50px;padding:5px;">URL</td>' +
		'<td style="width:200px;padding-top:5px;padding-bottom:5px;"><input type="text" id="hyperLink" value="http://" style="width:190px;border:1px solid #555555;background-color:#FFFFFF;"></td>' +
		'<tr><td style="padding:5px;">'+STR_LINK_TARGET+'</td>' +
		'<td style="padding-bottom:5px;"><select id="hyperLinkTarget"><option value="_blank" selected>'+STR_LINK_BLANK+'</option><option value="">'+STR_LINK_NOBLANK+'</option></select></td>' + 
		'<tr><td colspan="2" style="padding-bottom:5px;" align="center"><input type="button" name="button" value="'+STR_BUTTON_CONFIRM+'" ' +
		'onclick="javascript:HtmlDrawLinkEnd();"' +
		'style="border:1px solid #555555;"> <input type="button" name="button" value="'+STR_BUTTON_CANCEL+'" onclick="javascript:HtmlClearTemp();" style="border:1px solid #555555;"></td></tr>';
	str += '</table>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = 'Html_LINK';
}
function HtmlDrawLinkEnd()
{
	var range;
	var url = document.getElementById('hyperLink').value;
	var target = document.getElementById('hyperLinkTarget').value;
	if (url.match(/http:\/\/.{3,}/) == null) {
		alert(MSG_INPUT_URL);
		return false;
	}
	HtmlSelect();
	HtmlExecuteValue('CreateLink', url);
	var element;
    if (BROWSER == 'IE') {
		if (SELECTION.type.toLowerCase() == 'text') {
			element = RANGE.parentElement() ? RANGE.parentElement() : RANGE.item(0).parentElement();
		}
	} else {
		element = RANGE.startContainer.previousSibling;
    }
    if (element && target) {
        element.target = target;
    }
	HtmlClearTemp();
}
function HtmlDrawIcon()
{
	var top = HtmlGetTop('Html_ICON') + 32;
	var left = HtmlGetLeft('Html_ICON') + 1;
	var str = '';
	var iconNum = 24;
	str += '<table cellpadding="0" cellspacing="2" style="'+HtmlGetMenuCommonStyle(top, left)+'">';
	for (i = 0; i < iconNum; i++) {
		if (i == 0 || (i >= 6 && i%6 == 0)) {
			str += '<tr>';
		}
		var num;
		if ((i+1).toString(10).length < 2) {
			num = '0' + (i+1);
		} else {
			num = (i+1).toString(10);
		}
		var iconUrl = ICON_PATH + 'etc_' + num + '.gif';
		str += '<td style="padding:2px;border:0;cursor:pointer;" ' + 
		'onclick="javascript:HtmlExecute(\'Html_ICON_END\', \'' + iconUrl + '\');">' +
		'<img src="' + iconUrl + '" style="border:1px solid #EEEEEE;" onmouseover="javascript:this.style.borderColor=\'#AAAAAA\';" ' +
		'onmouseout="javascript:this.style.borderColor=\'#EEEEEE\';">' + '</td>';
		if (i >= 5 && i%(i-1) == 0) {
			str += '</tr>';
		}
	}
	str += '</table>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = 'Html_ICON';
}
function HtmlDrawSpecialchar()
{
	var top = HtmlGetTop('Html_SPECIALCHAR') + 32;
	var left = HtmlGetLeft('Html_SPECIALCHAR') + 1;
	var str = '';
	str += '<table cellpadding="0" cellspacing="2" style="'+HtmlGetMenuCommonStyle(top, left)+'">';
	for (i = 0; i < SPECIAL_CHARACTER.length; i++) {
		if (i == 0 || (i >= 10 && i%10 == 0)) {
			str += '<tr>';
		}
		str += '<td style="padding:2px;border:1px solid #AAAAAA;cursor:pointer;" ' + 
		'onclick="javascript:HtmlExecute(\'Html_SPECIALCHAR_END\', \'' + SPECIAL_CHARACTER[i] + '\');" ' +
		'onmouseover="javascript:this.style.borderColor=\'#000000\';" ' +
		'onmouseout="javascript:this.style.borderColor=\'#AAAAAA\';">' + SPECIAL_CHARACTER[i] + '</td>';
		if (i >= 9 && i%(i-1) == 0) {
			str += '</tr>';
		}
	}
	str += '</table>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = 'Html_SPECIALCHAR';
}
function HtmlDrawTableSelected(i, j)
{
	var text = i.toString(10) + ' by ' + j.toString(10) + ' Table';
	document.getElementById('tableLocation').innerHTML = text;
	var num = 10;
	for (m = 1; m <= num; m++) {
		for (n = 1; n <= num; n++) {
			var obj = document.getElementById('tableTd' + m.toString(10) + '_' + n.toString(10) + '');
			if (m <= i && n <= j) {
				obj.style.backgroundColor = '#AAAAAA';
			} else {
				obj.style.backgroundColor = '#FFFFFF';
			}
		}
	}
}
function HtmlDrawTable()
{
	var top = HtmlGetTop('Html_TABLE') + 32;
	var left = HtmlGetLeft('Html_TABLE') + 1;
	var str = '';
	var num = 10;
	str += '<table cellpadding="0" cellspacing="0" style="'+HtmlGetMenuCommonStyle(top, left)+'">';
	for (i = 1; i <= num; i++) {
		str += '<tr>';
		for (j = 1; j <= num; j++) {
			var value = i.toString(10) + ',' + j.toString(10);
			str += '<td id="tableTd' + i.toString(10) + '_' + j.toString(10) + 
			'" style="width:15px;height:15px;background-color:#FFFFFF;border:1px solid #DDDDDD;cursor:pointer;" ' + 
			'onclick="javascript:HtmlExecute(\'Html_TABLE_END\', \'' + value + '\');" ' +
			'onmouseover="javascript:HtmlDrawTableSelected(\''+i.toString(10)+'\', \''+j.toString(10)+'\');" ' + 
			'onmouseout="javascript:;">&nbsp;</td>';
		}
		str += '</tr>';
	}
	str += '<tr><td colspan="10" id="tableLocation" style="text-align:center;height:20px;"></td></tr>';
	str += '</table>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = 'Html_TABLE';
}
function HtmlDrawImage()
{
	var top = HtmlGetTop('Html_IMAGE') + 32;
	var left = HtmlGetLeft('Html_IMAGE') + 1;
	var str = '';
	str += '<div style="width:250px;'+HtmlGetMenuCommonStyle(top, left)+'">';
	str += '<iframe name="UploadIframe" id="UploadIframe" frameborder="0" style="width:250px;height:340px;padding:0;margin:0;border:0;">';
	str += '</iframe></div>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = 'Html_IMAGE';
	if (BROWSER == 'IE') {
		UPLOAD_DOCUMENT = UploadIframe.document;
	} else {
		UPLOAD_DOCUMENT = document.getElementById('UploadIframe').contentDocument;
	}
	str = '<div align="center">' +
		'<table cellpadding="0" cellspacing="0" style="width:100%;font-size:12px;">' + 
		'<tr><td colspan="2"><table border="0" style="margin-bottom:5px;"><tr><td id="imgPreview" style="width:240px;height:240px;border:1px solid #AAAAAA;background-color:#FFFFFF;" align="center" valign="middle">&nbsp;</td></tr></table></td></tr>' +  	
		'<tr><td style="width:50px;padding-left:5px;">'+STR_IMAGE_REMOTE+'</td>' +
		'<td style="width:200px;padding-bottom:5px;">' +
		'<input type="text" id="imgLink" value="http://" style="width:95%;border:1px solid #555555;display:block;" onchange="javascript:parent.HtmlImagePreview();">' +
		'</td>' +
		'<tr><td colspan="2" style="padding-bottom:5px;"><table border="0" style="width:100%;font-size:12px;"><tr>' +
		'<td width="10%" style="padding:5px;">'+STR_WIDTH+'</td><td width="23%"><input type="text" name="imgWidth" id="imgWidth" value="0" maxlength="4" style="width:40px;border:1px solid #555555;"></td>' +
		'<td width="10%" style="padding:5px;">'+STR_HEIGHT+'</td><td width="23%"><input type="text" name="imgHeight" id="imgHeight" value="0" maxlength="4" style="width:40px;border:1px solid #555555;"></td>' +
		'<td width="10%" style="padding:5px;">'+STR_BORDER+'</td><td width="23%"><input type="text" name="imgBorder" id="imgBorder" value="0" maxlength="1" style="width:40px;border:1px solid #555555;"></td></tr></table></td></tr>' +  	
		'<tr><td colspan="2" style="margin:5px;padding-bottom:5px;" align="center">' +
		'<input type="button" name="button" value="'+STR_BUTTON_PREVIEW+'" onclick="javascript:parent.HtmlImagePreview();" style="border:1px solid #555555;"> ' +
		'<input type="button" name="button" value="'+STR_BUTTON_CONFIRM+'" onclick="javascript:parent.HtmlDrawImageEnd();" style="border:1px solid #555555;"> ' +
		'<input type="button" name="button" value="'+STR_BUTTON_CANCEL+'" onclick="javascript:parent.HtmlClearTemp();" style="border:1px solid #555555;"></td></tr>' + 
		'</table></div>';
	UPLOAD_DOCUMENT.open();
	UPLOAD_DOCUMENT.write(str);
	UPLOAD_DOCUMENT.close();
	UPLOAD_DOCUMENT.body.style.color = MENU_TEXT_COLOR;
	UPLOAD_DOCUMENT.body.style.backgroundColor = MENU_BG_COLOR;
	UPLOAD_DOCUMENT.body.style.margin = 0;
	UPLOAD_DOCUMENT.body.scroll = 'no';
}


function HtmlImagePreview()
{
	var url = UPLOAD_DOCUMENT.getElementById('imgLink').value;
	if (HtmlCheckImageFileType(url, "/") == false)
		{
			return false;
		}
	var imgObj = UPLOAD_DOCUMENT.createElement("IMG");
	imgObj.src = url;
	var width = parseInt(imgObj.width);
	var height = parseInt(imgObj.height);
	UPLOAD_DOCUMENT.getElementById('imgWidth').value = width;
	UPLOAD_DOCUMENT.getElementById('imgHeight').value = height;
	var rate = parseInt(width/height);
	if (width >230 && height <= 230) {
		width = 230;
		height = parseInt(width/rate);
	} else if (width <=230 && height > 230) {
		height = 230;
		width = parseInt(height*rate);
	} else if (width >230 && height > 230) {
		if (width >= height) {
			width = 230;
			height = parseInt(width/rate);
		} else {
			height = 230;
			width = parseInt(height*rate);
		}
	}
	imgObj.style.width = width;
	imgObj.style.height = height;
	var el = UPLOAD_DOCUMENT.getElementById('imgPreview');
	if (el.hasChildNodes()) {
		el.removeChild(el.childNodes[0]);
	}
	el.appendChild(imgObj);
	return imgObj;
}
function HtmlDrawImageEnd()
{
	var url = UPLOAD_DOCUMENT.getElementById('imgLink').value;
	var width = UPLOAD_DOCUMENT.getElementById('imgWidth').value;
	var height = UPLOAD_DOCUMENT.getElementById('imgHeight').value;
	var border = UPLOAD_DOCUMENT.getElementById('imgBorder').value;

	if (HtmlCheckImageFileType(url, "/") == false) {
		return false;
	}

	if (width.match(/^\d+$/) == null) {
		alert(MSG_INVALID_WIDTH);
		return false;
	}
	if (height.match(/^\d+$/) == null) {
		alert(MSG_INVALID_HEIGHT);
		return false;
	}
	if (border.match(/^\d+$/) == null) {
		alert(MSG_INVALID_BORDER);
		return false;
	}
		if (HtmlCheckImageFileType(url, "/") == false) {
			return false;
		}
		HtmlInsertImage(url, width, height, border);
}
function HtmlInsertImage(url, width, height, border)
{
	var element = document.createElement("img");
	element.src = url;
	if (width > 0) {
		element.style.width = width;
	}
	if (height > 0) {
		element.style.height = height;
	}
	element.border = border;
	HtmlSelect();
	HtmlInsertItem(element);
	HtmlClearTemp();
}
function HtmlDrawMedia()
{
	var top = HtmlGetTop('Html_MEDIA') + 32;
	var left = HtmlGetLeft('Html_MEDIA') + 1;
	var str = '';
	var str = '';
	str += '' +
		'<table cellpadding="0" cellspacing="0" style="width:250px;'+HtmlGetMenuCommonStyle(top, left)+'">' + 
		'<tr><td colspan="2"><table border="0"><tr><td id="mediaPreview" style="width:240px;height:240px;border:1px solid #AAAAAA;background-color:#FFFFFF;" align="center" valign="middle">&nbsp;</td></tr></table></td></tr>' +  	
		'<tr><td style="width:50px;padding:5px;">'+STR_IMAGE_REMOTE+'</td>' +
		'<td style="width:200px;padding-bottom:5px;"><input type="text" id="mediaLink" value="http://" style="width:190px;border:1px solid #555555;" onchange="javascript:HtmlMediaPreview();"></td>' +
		'<tr><td colspan="2" style="margin:5px;padding-bottom:5px;" align="center">' +
		'<input type="button" name="button" value="'+STR_BUTTON_LISTENING+'" onclick="javascript:HtmlMediaPreview();" style="border:1px solid #555555;"> ' +
		'<input type="button" name="button" value="'+STR_BUTTON_CONFIRM+'" onclick="javascript:HtmlDrawMediaEnd();" style="border:1px solid #555555;"> ' +
		'<input type="button" name="button" value="'+STR_BUTTON_CANCEL+'" onclick="javascript:HtmlClearTemp();" style="border:1px solid #555555;"></td></tr>' + 
		'</table>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = 'Html_MEDIA';
}
function HtmlGetMediaHtmlTag(url, width, height)
{
	var element = document.createElement("embed");
	element.src = url;
	element.width = 230;
	element.height = 230;
	element.loop = "true";
	element.autostart = "true";
	return element;
}
function HtmlMediaPreview()
{
	var url = document.getElementById('mediaLink').value;
	if (HtmlCheckMediaFileType(url, "/") == false) {
		return false;
	}
	var el = document.getElementById('mediaPreview');
	if (el.hasChildNodes()) {
		el.removeChild(el.childNodes[0]);
	}
	el.appendChild(HtmlGetMediaHtmlTag(url, '230', '230'));
}
function HtmlDrawMediaEnd()
{
	var url = document.getElementById('mediaLink').value;
	if (HtmlCheckMediaFileType(url, "/") == false) {
		return false;
	}
	HtmlSelect();
	HtmlInsertItem(HtmlGetMediaHtmlTag(url, '100', '100'));
	HtmlClearTemp();
}
function HtmlDrawFlash()
{
	var top = HtmlGetTop('Html_FLASH') + 32;
	var left = HtmlGetLeft('Html_FLASH') + 1;
	var str = '';
	str +=''+
		'<table cellpadding="0" cellspacing="0" style="width:250px;'+HtmlGetMenuCommonStyle(top, left)+'">' + 
		'<tr><td colspan="2"><table border="0"><tr><td id="flashPreview" style="width:240px;height:240px;border:1px solid #AAAAAA;background-color:#FFFFFF;" align="center" valign="middle">&nbsp;</td></tr></table></td></tr>' +  	
		'<tr><td style="width:50px;padding:5px;">'+STR_IMAGE_REMOTE+'</td>' +
		'<td style="width:200px;padding-bottom:5px;"><input type="text" id="flashLink" value="http://" style="width:190px;border:1px solid #555555;" onchange="javascript:HtmlFlashPreview();"></td>' +
		'<tr><td colspan="2" style="margin:5px;padding-bottom:5px;" align="center">' +
		'<input type="button" name="button" value="'+STR_BUTTON_PREVIEW+'" onclick="javascript:HtmlFlashPreview();" style="border:1px solid #555555;"> ' +
		'<input type="button" name="button" value="'+STR_BUTTON_CONFIRM+'" onclick="javascript:HtmlDrawFlashEnd();" style="border:1px solid #555555;"> ' +
		'<input type="button" name="button" value="'+STR_BUTTON_CANCEL+'" onclick="javascript:HtmlClearTemp();" style="border:1px solid #555555;"></td></tr>' + 
		'</table>';
	document.getElementById('popupData').innerHTML = str;
	document.getElementById('popupName').innerHTML = 'Html_FLASH';
}
function HtmlGetFlashHtmlTag(url, width, height)
{
	var element = document.createElement("embed");
	element.src = url;
	element.quality = "high";
	element.type = "application/x-shockwave-flash";
	element.width = width;
	element.height = height;
	return element;
}
function HtmlFlashPreview()
{
	var url = document.getElementById('flashLink').value;
	if (HtmlCheckFlashFileType(url, "/") == false) {
		return false;
	}
	var el = document.getElementById('flashPreview');
	if (el.hasChildNodes()) {
		el.removeChild(el.childNodes[0]);
	}
	el.appendChild(HtmlGetFlashHtmlTag(url, '230', '230'));
}
function HtmlDrawFlashEnd()
{
	var url = document.getElementById('flashLink').value;
	if (HtmlCheckFlashFileType(url, "/") == false) {
		return false;
	}
	HtmlSelect();
	HtmlInsertItem(HtmlGetFlashHtmlTag(url, '450', '350'));
	HtmlClearTemp();
}

function HtmlSelection()
{
	if (BROWSER == 'IE') {
		SELECTION = EDITFORM_DOCUMENT.selection;
		RANGE = SELECTION.createRange();
		RANGE_TEXT = RANGE.text;
	} else {
		SELECTION = document.getElementById("EditForm").contentWindow.getSelection();
        RANGE = SELECTION.getRangeAt(0);
		RANGE_TEXT = RANGE.toString();
	}
}
function HtmlSelect()
{
	if (BROWSER == 'IE') {
		RANGE.select();
	}
}
function HtmlInsertItem(insertNode)
{
	if (BROWSER == 'IE') {
		RANGE.parentElement() ? RANGE.pasteHTML(insertNode.outerHTML) : RANGE.item(0).outerHTML = insertNode.outerHTML;
	} else {
        SELECTION.removeAllRanges();
		RANGE.deleteContents();
        var startRangeNode = RANGE.startContainer;
        var startRangeOffset = RANGE.startOffset;
        var newRange = document.createRange();
		if (startRangeNode.nodeType == 3 && insertNode.nodeType == 3) {
            startRangeNode.insertData(startRangeOffset, insertNode.nodeValue);
            newRange.setEnd(startRangeNode, startRangeOffset + insertNode.length);
            newRange.setStart(startRangeNode, startRangeOffset + insertNode.length);
        } else {
            var afterNode;
            if (startRangeNode.nodeType == 3) {
                var textNode = startRangeNode;
                startRangeNode = textNode.parentNode;
                var text = textNode.nodeValue;
                var textBefore = text.substr(0, startRangeOffset);
                var textAfter = text.substr(startRangeOffset);
                var beforeNode = document.createTextNode(textBefore);
                var afterNode = document.createTextNode(textAfter);
                startRangeNode.insertBefore(afterNode, textNode);
                startRangeNode.insertBefore(insertNode, afterNode);
                startRangeNode.insertBefore(beforeNode, insertNode);
                startRangeNode.removeChild(textNode);
            } else {
				if (startRangeNode.tagName.toLowerCase() == 'html') {
					startRangeNode = startRangeNode.childNodes[0].nextSibling;
					afterNode = startRangeNode.childNodes[0];
				} else {
					afterNode = startRangeNode.childNodes[startRangeOffset];
				}
				startRangeNode.insertBefore(insertNode, afterNode);
            }
            newRange.setEnd(afterNode, 0);
            newRange.setStart(afterNode, 0);
        }
        SELECTION.addRange(newRange);
	}
}
function HtmlExecuteValue(cmd, value)
{
	if (BROWSER == 'IE') {
		RANGE.execCommand(cmd, false, value);
	} else {
		EDITFORM_DOCUMENT.execCommand(cmd, false, value);
	}
}
function HtmlSimpleExecute(cmd)
{
	EDITFORM_DOCUMENT.execCommand(cmd, false, null);
	HtmlClearTemp();
	EditForm.focus();
}
function HtmlExecute(mode, value)
{
	switch (mode)
	{
		case 'Html_SOURCE':
			var length = document.getElementById(TOP_TOOLBAR_ICON[0][0]).src.length - 10;
			var image = document.getElementById(TOP_TOOLBAR_ICON[0][0]).src.substr(length,10);
			if (image == 'source.gif') {
				document.getElementById("CodeForm").value = HtmlHtmlToXhtml(EDITFORM_DOCUMENT.body.innerHTML);
				document.getElementById("editIframe").style.display = 'none';
				document.getElementById("editTextarea").style.display = 'block';
				HtmlDisableToolbar(true);
			} else {
				EDITFORM_DOCUMENT.body.innerHTML = document.getElementById("CodeForm").value;
				document.getElementById("editTextarea").style.display = 'none';
				document.getElementById("editIframe").style.display = 'block';
				HtmlDisableToolbar(false);
			}
			HtmlClearTemp();
			break;
		case 'Html_COPY':
			HtmlSimpleExecute('copy');
			break;
		case 'Html_PASTE':
			HtmlSimpleExecute('paste');
			break;
		case 'Html_SELECTALL':
			HtmlSimpleExecute('selectall');
			break;
		case 'Html_SUBSCRIPT':
			HtmlSimpleExecute('subscript');
			break;
		case 'Html_SUPERSCRIPT':
			HtmlSimpleExecute('superscript');
			break;
		case 'Html_BOLD':
			HtmlSimpleExecute('bold');
			break;
		case 'Html_ITALIC':
			HtmlSimpleExecute('italic');
			break;
		case 'Html_UNDERLINE':
			HtmlSimpleExecute('underline');
			break;
		case 'Html_STRIKE':
			HtmlSimpleExecute('strikethrough');
			break;
		case 'Html_JUSTIFYLEFT':
			HtmlSimpleExecute('justifyleft');
			break;
		case 'Html_JUSTIFYCENTER':
			HtmlSimpleExecute('justifycenter');
			break;
		case 'Html_JUSTIFYRIGHT':
			HtmlSimpleExecute('justifyright');
			break;
		case 'Html_JUSTIFYFULL':
			HtmlSimpleExecute('justifyfull');
			break;
		case 'Html_NUMBEREDLIST':
			HtmlSimpleExecute('insertorderedlist');
			break;
		case 'Html_UNORDERLIST':
			HtmlSimpleExecute('insertunorderedlist');
			break;
		case 'Html_INDENT':
			HtmlSimpleExecute('indent');
			break;
		case 'Html_OUTDENT':
			HtmlSimpleExecute('outdent');
			break;
		case 'Html_REMOVE':
			HtmlSimpleExecute('removeformat');
			break;
		case 'Html_ZOOM':
			EditForm.focus();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawZoom();
			}
			break;
		case 'Html_ZOOM_END':
			EditForm.focus();
			EDITFORM_DOCUMENT.body.style.zoom = value;
			HtmlClearTemp();
			break;
		case 'Html_TITLE':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawTitle();
			}
			break;
		case 'Html_TITLE_END':
			EditForm.focus();
			value = '<' + value + '>';
			HtmlSelect();
			HtmlExecuteValue('FormatBlock', value);
			HtmlClearTemp();
			break;
		case 'Html_FONTNAME':
			EditForm.focus();
			HtmlSelection();
			if (HtmlCheckSelected() == false) {
				return false;
			}
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawFontname();
			}
			break;
		case 'Html_FONTNAME_END':
			EditForm.focus();
			HtmlSelect();
			HtmlExecuteValue('fontname', value);
			HtmlClearTemp();
			break;
		case 'Html_FONTSIZE':
			EditForm.focus();
			HtmlSelection();
			if (HtmlCheckSelected() == false) {
				return false;
			}
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawFontsize();
			}
			break;
		case 'Html_FONTSIZE_END':
			EditForm.focus();
			value = value.substr(0, 1);
			HtmlSelect();
			HtmlExecuteValue('fontsize', value);
			HtmlClearTemp();
			break;
		case 'Html_TEXTCOLOR':
			EditForm.focus();
			HtmlSelection();
			if (HtmlCheckSelected() == false) {
				return false;
			}
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawColorTable('Html_TEXTCOLOR');
			}
			break;
		case 'Html_TEXTCOLOR_END':
			EditForm.focus();
			HtmlSelect();
			HtmlExecuteValue('ForeColor', value);
			HtmlClearTemp();
			break;
		case 'Html_BGCOLOR':
			EditForm.focus();
			HtmlSelection();
			if (HtmlCheckSelected() == false) {
				return false;
			}
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawColorTable('Html_BGCOLOR');
			}
			break;
		case 'Html_BGCOLOR_END':
			EditForm.focus();
			if (BROWSER == 'IE') {
				HtmlSelect();
				HtmlExecuteValue('BackColor', value);
			} else {
				var startRangeNode = RANGE.startContainer;
				if (RANGE.toString() != '' && startRangeNode.nodeType == 3) {
					var parent = startRangeNode.parentNode;
					var element = document.createElement("font");
					element.style.backgroundColor = value;
					element.appendChild(RANGE.extractContents());
					var startRangeOffset = RANGE.startOffset;
					var newRange = document.createRange();
					var afterNode;
					var textNode = startRangeNode;
					startRangeNode = textNode.parentNode;
					var text = textNode.nodeValue;
					var textBefore = text.substr(0, startRangeOffset);
					var textAfter = text.substr(startRangeOffset);
					var beforeNode = document.createTextNode(textBefore);
					var afterNode = document.createTextNode(textAfter);
					startRangeNode.insertBefore(afterNode, textNode);
					startRangeNode.insertBefore(element, afterNode);
					startRangeNode.insertBefore(beforeNode, element);
					startRangeNode.removeChild(textNode);
					newRange.setEnd(afterNode, 0);
					newRange.setStart(afterNode, 0);
					SELECTION.addRange(newRange);
				}
			}
			HtmlClearTemp();
			break;
		case 'Html_LINK':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawLink();
			}
			break;
		case 'Html_UNLINK':
			HtmlSimpleExecute('unlink');
			break;
		case 'Html_ICON':
			EditForm.focus();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawIcon();
			}
			break;
		case 'Html_ICON_END':
			EditForm.focus();
			EDITFORM_DOCUMENT.execCommand('InsertImage', false, value);
			HtmlClearTemp();
			break;
		case 'Html_IMAGE':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawImage();
			}
			break;
		case 'Html_MEDIA':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawMedia();
			}
			break;
		case 'Html_FLASH':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawFlash();
			}
			break;
		case 'Html_SPECIALCHAR':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawSpecialchar();
			}
			break;
		case 'Html_SPECIALCHAR_END':
			EditForm.focus();
			HtmlSelect();
			var element = document.createElement("span");
			element.appendChild(document.createTextNode(value));
			HtmlInsertItem(element);
			HtmlClearTemp();
			break;
		case 'Html_LAYER':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawColorTable('Html_LAYER');
			}
			break;
		case 'Html_LAYER_END':
			EditForm.focus();
			var element = document.createElement("div");
			element.style.width = "100px";
			element.style.height = "100px";
			element.style.padding = "5px";
			element.style.border = "1px solid #AAAAAA";
			element.style.backgroundColor = value;
			element.innerHTML = "&nbsp;";
			HtmlSelect();
			HtmlInsertItem(element);
			HtmlClearTemp();
			break;
		case 'Html_TABLE':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawTable();
			}
			break;
		case 'Html_TABLE_END':
			var location = value.split(',');
			var element = document.createElement("table");
			element.cellPadding = 0;
			element.cellSpacing = 0;
			element.border = 1;
			element.style.width = "100px";
			element.style.height = "100px";
			for (var i = 0; i < location[0]; i++) {
				var rowElement = element.insertRow(i);
				for (var j = 0; j < location[1]; j++) {
					var cellElement = rowElement.insertCell(j);
					cellElement.innerHTML = "&nbsp;";
				}
			}
			EditForm.focus();
			HtmlSelect();
			HtmlInsertItem(element);
			HtmlClearTemp();
			break;
		case 'Html_HR':
			EditForm.focus();
			HtmlSelection();
			if (document.getElementById('popupName').innerHTML == mode) {
				HtmlClearTemp();
			} else {
				HtmlDrawColorTable('Html_HR');
			}
			break;
		case 'Html_HR_END':
			EditForm.focus();
			var element = document.createElement("hr");
			element.width = "100%";
			element.color = value;
			element.size = 1;
			HtmlSelect();
			HtmlInsertItem(element);
			HtmlClearTemp();
			break;
		case 'Html_DATE':
			EditForm.focus();
			HtmlSelection();
			var date = new Date();
			var year = date.getFullYear().toString(10);
			var month = (date.getMonth() + 1).toString(10);
			month = month.length < 2 ? '0' + month : month;
			var day = date.getDate().toString(10);
			day = day.length < 2 ? '0' + day : day;
			var value = year + '-' + month + '-' + day;
			var element = document.createElement("span");
			element.appendChild(document.createTextNode(value));
			HtmlInsertItem(element);
			HtmlClearTemp();
			break;
		case 'Html_TIME':
			EditForm.focus();
			HtmlSelection();
			var date = new Date();
			var hour = date.getHours().toString(10);
			hour = hour.length < 2 ? '0' + hour : hour;
			var minute = date.getMinutes().toString(10);
			minute = minute.length < 2 ? '0' + minute : minute;
			var second = date.getSeconds().toString(10);
			second = second.length < 2 ? '0' + second : second;
			var value = hour + ':' + minute + ':' + second;
			var element = document.createElement("span");
			element.appendChild(document.createTextNode(value));
			HtmlInsertItem(element);
			HtmlClearTemp();
			break;
		default: 
			break;
	}
}
function HtmlOverIcon(obj)
{
	obj.style.border = '1px solid ' + MENU_BORDER_COLOR;
}
function HtmlOutIcon(obj)
{
	obj.style.border = '1px solid ' + TOOLBAR_BG_COLOR;
}
function HtmlDisableToolbar(flag)
{
	if (flag == true) {
		document.getElementById(TOP_TOOLBAR_ICON[0][0]).src = IMAGE_PATH + 'design.gif';
		for (i = 1; i < TOOLBAR_ICON.length; i++) {
			var el = document.getElementById(TOOLBAR_ICON[i][0]);
			el.style.visibility = 'hidden';
		}
	} else {
		document.getElementById(TOP_TOOLBAR_ICON[0][0]).src = IMAGE_PATH + 'source.gif';
		for (i = 0; i < TOOLBAR_ICON.length; i++) {
			var el = document.getElementById(TOOLBAR_ICON[i][0]);
			el.style.visibility = 'visible';
			EDITFORM_DOCUMENT.designMode = 'On';
		}
	}
}
function HtmlWriteFullHtml(documentObj)
{
	var editHtmlData = '<html>\r\n<head>\r\n<title>htmlEditor</title>\r\n<style type="text/css">\r\np {margin:0;}\r\n</style>\r\n</head>\r\n';
	editHtmlData += '<body style="font-size:12px;font-family:'+EDITOR_FONT_FAMILY+';margin:2px;background-color:' + FORM_BG_COLOR + '">\r\n';
	editHtmlData += HtmlHtmlentitiesDecode(document.getElementsByName(eval(OBJ_NAME).hiddenName)[0].value);
	editHtmlData += '\r\n</body>\r\n</html>\r\n';
	documentObj.open();
	documentObj.write(editHtmlData);
	documentObj.close();
}



function htmlEditor(objName) 
{
	this.objName = objName;
	this.hiddenName = objName;
	this.width = "100%";
	this.height = "200px";
	this.imagePath = 'images/';
	this.iconPath = 'images/face/';
	this.imageAttachPath = 'attached/';
	this.imageUploadCgi = "upload.php";
	this.menuBorderColor = '#4169e1';
	this.menuBgColor = '#e6e6fa';
	this.menuTextColor = '#8b0000';
	this.menuSelectedColor = '#6495ed';
	this.toolbarBgColor = '#DDDDDD';
	this.formBorderColor = '#AAAAAA';
	this.formBgColor = '#FFFFFF';

	this.init = function()
	{
		IMAGE_PATH = this.imagePath;
		ICON_PATH = this.iconPath;
		IMAGE_ATTACH_PATH = this.imageAttachPath;
		IMAGE_UPLOAD_CGI = this.imageUploadCgi;
		MENU_BORDER_COLOR = this.menuBorderColor;
		MENU_BG_COLOR = this.menuBgColor;
		MENU_TEXT_COLOR = this.menuTextColor;
		MENU_SELECTED_COLOR = this.menuSelectedColor;
		TOOLBAR_BG_COLOR = this.toolbarBgColor;
		FORM_BORDER_COLOR = this.formBorderColor;
		FORM_BG_COLOR = this.formBgColor;
		OBJ_NAME = this.objName;
		BROWSER = HtmlGetBrowser();
		TOOLBAR_ICON = Array();
		for (var i = 0; i < TOP_TOOLBAR_ICON.length; i++) {
			TOOLBAR_ICON.push(TOP_TOOLBAR_ICON[i]);
		}
		for (var i = 0; i < BOTTOM_TOOLBAR_ICON.length; i++) {
			TOOLBAR_ICON.push(BOTTOM_TOOLBAR_ICON[i]);
		}
	}
	this.show = function()
	{
		this.init();
		var iframeSize = '';
		iframeSize += 'width:' + this.width + ';';
		iframeSize += 'height:' + this.height + ';';
		if (BROWSER == '') {
			var htmlData = '<div id="editTextarea" style="' + iframeSize + '">' +
			'<textarea name="CodeForm" id="CodeForm" style="' + iframeSize + 
			'padding:2px;margin:0;border:1px solid '+ FORM_BORDER_COLOR + 
			';font-size:12px;line-height:16px;font-family:'+EDITOR_FONT_FAMILY+';background-color:'+ 
			FORM_BG_COLOR +';">' + document.getElementsByName(this.hiddenName)[0].value + '</textarea></div>';
			document.open();
			document.write(htmlData);
			document.close();
			return;
		}
		var htmlData = '<div id="Htmldiv" style="font-family:'+EDITOR_FONT_FAMILY+';">' +
			'<div id="htmltoolbar" style="padding:1px;background-color:'+ TOOLBAR_BG_COLOR +'">' +
			'<table cellpadding="0" cellspacing="0" border="0"><tr>';
		for (i = 0; i < TOP_TOOLBAR_ICON.length; i++) {
			htmlData += '<td width="' + TOP_TOOLBAR_ICON[i][2] + 
			'" height="' + TOP_TOOLBAR_ICON[i][3] + '"><img id="'+ TOP_TOOLBAR_ICON[i][0] +'" src="' + IMAGE_PATH + 
			TOP_TOOLBAR_ICON[i][1] + '" width="' + TOP_TOOLBAR_ICON[i][2] + 
			'" height="' + TOP_TOOLBAR_ICON[i][3] + '" alt="' + TOP_TOOLBAR_ICON[i][4] + 
			'" title="' + TOP_TOOLBAR_ICON[i][4] + 
			'" align="absmiddle" style="border:1px solid '+ 
			TOOLBAR_BG_COLOR +';cursor:pointer;" onclick="javascript:HtmlExecute(\''+ TOP_TOOLBAR_ICON[i][0] +'\');"'+
			' onmouseover="javascript:HtmlOverIcon(this);" onmouseout="javascript:HtmlOutIcon(this);"></td>';
		}
		htmlData += '</tr></table><table cellpadding="0" cellspacing="0" border="0"><tr>';
		for (i = 0; i < BOTTOM_TOOLBAR_ICON.length; i++) {
			htmlData += '<td width="' + BOTTOM_TOOLBAR_ICON[i][2] + 
			'" height="' + BOTTOM_TOOLBAR_ICON[i][3] + '"><img id="'+ BOTTOM_TOOLBAR_ICON[i][0] +'" src="' + IMAGE_PATH + 
			BOTTOM_TOOLBAR_ICON[i][1] + '" width="' + BOTTOM_TOOLBAR_ICON[i][2] + 
			'" height="' + BOTTOM_TOOLBAR_ICON[i][3] + '" alt="' + BOTTOM_TOOLBAR_ICON[i][4] + 
			'" title="' + BOTTOM_TOOLBAR_ICON[i][4] + 
			'" align="absmiddle" style="border:1px solid '+ 
			TOOLBAR_BG_COLOR +';cursor:pointer;" onclick="javascript:HtmlExecute(\''+ BOTTOM_TOOLBAR_ICON[i][0] +'\');"'+
			' onmouseover="javascript:HtmlOverIcon(this);" onmouseout="javascript:HtmlOutIcon(this);"></td>';
		}
		htmlData += '</tr>' +
			'</table>' +
			'</div>' +
			'<div id="editIframe">' +
			'<iframe name="EditForm" id="EditForm" frameborder="0" style="' + iframeSize + 
			'padding:0;margin:0;border:1px solid '+ FORM_BORDER_COLOR +'">' +
			'</iframe>' +
			'</div>' +
			'<div id="editTextarea" style="display:none;">' +
			'<textarea onkeydown="javascript:if(event.ctrlKey && event.keyCode==13){submitform();return false}" name="CodeForm" id="CodeForm" style="' + iframeSize + 
			'padding:2px;margin:0;border:1px solid '+ FORM_BORDER_COLOR + 
			';font-size:12px;line-height:16px;font-family:'+EDITOR_FONT_FAMILY+';background-color:'+ 
			FORM_BG_COLOR +';"></textarea>' +
			'</div>' +
			'</div>' +
			'<span id="popupName" style="display:none;"></span>' +
			'<span id="popupData"></span>';
		document.open();
		document.write(htmlData);
		document.close();
		if (BROWSER == 'IE') {
			EDITFORM_DOCUMENT = document.frames("EditForm").document;
		} else {
			EDITFORM_DOCUMENT = document.getElementById('EditForm').contentDocument;
		}
		EDITFORM_DOCUMENT.designMode = 'On';
		HtmlWriteFullHtml(EDITFORM_DOCUMENT);
		var el = EDITFORM_DOCUMENT.body;
		if (BROWSER == 'OPERA') {
			el.onclick = HtmlClearTemp;
		} else {
			if (el.addEventListener){
				el.addEventListener('click', HtmlClearTemp, false); 
			} else if (el.attachEvent){
				el.attachEvent('onclick', HtmlClearTemp);
			}
		}
		el.onkeydown=new Function("with(EditForm.event)if(ctrlKey && keyCode==13){submitform();return false}")

	}
	this.data = function()
	{
		if (BROWSER == '') {
			htmlResult = document.getElementById("CodeForm").value;
			htmlResult=cleanHtml(htmlResult);
			htmlResult=FilterScript(htmlResult);
			document.getElementsByName(this.hiddenName)[0].value = htmlResult;
			return htmlResult;
		}
		var length = document.getElementById(TOP_TOOLBAR_ICON[0][0]).src.length - 10;
		var image = document.getElementById(TOP_TOOLBAR_ICON[0][0]).src.substr(length,10);
		var htmlResult;
		if (image == 'source.gif') {
			htmlResult = HtmlHtmlToXhtml(EDITFORM_DOCUMENT.body.innerHTML);
		} else {
			htmlResult = document.getElementById("CodeForm").value;
		}		
		htmlResult=cleanHtml(htmlResult);
		htmlResult=FilterScript(htmlResult);
		document.getElementsByName(this.hiddenName)[0].value = htmlResult;
		return htmlResult;
	}
}

var strlength;

function CopyBody()
{
EditForm.focus();
EditForm.document.execCommand('selectAll');
EditForm.document.execCommand('copy');
//alert("你发表的内容已被复制到剪贴板，如果发帖不成功,请在编辑框中按下CTRL+V即可找回帖子内容!")
}

//提交表单
function submitform(){
   	if(htmlsubmit()==true)CopyBody();document.topicform.submit();
}

//检测表单
function htmlsubmit() {
        var content = editor.data();
		strlength=document.getElementsByName("content").item(0).value.length;
		if (strlength>25600||strlength<5){
			alert("您输入的文章长度为"+strlength+"，长度必须大于5且小于25600，请修正之后再继续。");
			return false;
		}
		else if(document.getElementsByName("Caption").item(0).value==""){
			alert("标题不能为空。");
			document.getElementById("Caption").style.backgroundColor="#FFEEEE";
			return false;
		}
		else{ 
			return true;
		}
}

function autosave(){
	if(document.getElementsByName("usereditor").item(0).checked==true){
		var content = editor.data();
	}
	else if(document.getElementsByName("usereditor").item(1).checked==true){
		document.getElementsByName("content").item(0).value = document.getElementById("CodeForm").value;
	}
	set_cookie("tempcontent",document.getElementsByName("content").item(0).value);
}


setInterval("autosave()",5000);

function presskey(eventobject)
{
var eve=eventobject||window.event;
   if(eve.ctrlKey && eve.keyCode==13){submitform();}else{return false}
}

function Especial(A,B){
fontbegin=A
fontend=B
fontchuli();
}
function fontchuli(){
if(ns6){
	sel=EditForm.getSelection()
	var html = EditForm.document.body.innerHTML;
	if (sel!=""){
		html = html.replace(sel,fontbegin + sel + fontend);
		EditForm.document.body.innerHTML=html}
	else
		{EditForm.document.body.innerHTML+=(fontbegin +"" +  fontend)}
	EditForm.focus()
}else{
	EditForm.focus()
	sel=EditForm.document.selection.createRange();
	if (sel.text!=""){sel.pasteHTML(fontbegin + sel.text + fontend);}
	else{sel.pasteHTML(fontbegin + "" + fontend);}
}
}

function preview()
{
if(htmlsubmit()){
document.form1.Caption.value=document.topicform.Caption.value;
document.form1.content.value=document.topicform.content.value;
var popupWin = window.open('see.asp?action=preview','showgg','width=500,height=400,resizable=1,scrollbars=yes,menubar=no,status=yes');
document.form1.submit()
}
}

function Goreset(){EditForm.document.body.innerHTML=document.topicform.content.value;}

function DoTitle(addTitle){
document.getElementById("caption").value=addTitle+document.getElementById("caption").value;
document.getElementById("caption").focus();
document.getElementById("caption").value+="";
}

//投票
function SetNum(obj){
	str = '';
	if(!obj.value){
		obj.value = 1;
	}
	for(i=1;i<=obj.value;i++){
		str += '<div>选项'+i+'：<input type="text" name="votes'+i+'" style="width:70%" \/><\/div>';
	}
	document.getElementById("optionid").innerHTML=str+'<input type="hidden" name="autovalue" value="'+obj.value+'" \/>';
}

