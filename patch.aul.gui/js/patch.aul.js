$(function(){

function setBoolValue(selector, value)
{
	$(selector).prop('checked', value);
}

function setIntValue(selector, value)
{
	$(selector).prop('value', value);
}

function setFloatValue(selector, value)
{
	$(selector).prop('value', value);
}

function setColorValue(selector, value)
{
	$(selector).spectrum('set', value);
}

function getBoolValue(selector)
{
	return $(selector).prop('checked');
}

function getIntValue(selector)
{
	return $(selector).prop('value') - 0;
}

function getFloatValue(selector)
{
	return $(selector).prop('value') - 0;
}

function getPureColorValue(selector)
{
	return $(selector).spectrum('get').toHexString() + '';
}

function getColorValue(selector)
{
	return $(selector).spectrum('get').toHex() + '';
}

function getRGBColorValue(selector)
{
	return $(selector).spectrum('get').toRgb();
}

function loadFile(showDialog)
{
	var fileName = 'patch.aul.json';

	if (showDialog)
	{
		var shell = new ActiveXObject('WScript.Shell');
		var cd = shell.currentDirectory;

		var fso = new ActiveXObject("Scripting.FileSystemObject");
		var path = window.location.pathname.replace('/', '');
		path = fso.GetParentFolderName(path);
		path = '"' + path + '\\file.dialog.exe" -Open ' + '"' + cd + '"';

		var exec = shell.Exec(path);
		fileName = exec.StdOut.ReadLine();

		if (!fileName)
			return;
	}

	try {
		var fs = new ActiveXObject('Scripting.FileSystemObject');
		var file = fs.OpenTextFile(fileName, 1, false);
		var to_json = file.ReadAll(); 
		file.Close();
		var json = JSON.parse(to_json);
	} catch (e) {
		alert('Error');
		return;
	}

	if (json.console)
	{
		// console

		setBoolValue('#json #console #visible', json.console.visible);

		if (json.console.rect)
		{
			setIntValue('#json #console #rect_0', json.console.rect[0]);
			setIntValue('#json #console #rect_1', json.console.rect[1]);
			setIntValue('#json #console #rect_2', json.console.rect[2]);
			setIntValue('#json #console #rect_3', json.console.rect[3]);
		}
	}

	if (json.theme_cc)
	{
		if (json.theme_cc.layer)
		{
			// theme_cc layer

			setIntValue('#json #theme_cc #layer #height_large', json.theme_cc.layer.height_large);
			setIntValue('#json #theme_cc #layer #height_medium', json.theme_cc.layer.height_medium);
			setIntValue('#json #theme_cc #layer #height_small', json.theme_cc.layer.height_small);

			setColorValue('#json #theme_cc #layer #link_col', json.theme_cc.layer.link_col);
			setColorValue('#json #theme_cc #layer #clipping_col', json.theme_cc.layer.clipping_col);
			setColorValue('#json #theme_cc #layer #lock_col_0', json.theme_cc.layer.lock_col[0]);
			setColorValue('#json #theme_cc #layer #lock_col_1', json.theme_cc.layer.lock_col[1]);

			setFloatValue('#json #theme_cc #layer #hide_alpha', json.theme_cc.layer.hide_alpha);
		}

		if (json.theme_cc.object)
		{
			// theme_cc object

			setColorValue('#json #theme_cc #object #media_col_0', json.theme_cc.object.media_col[0]);
			setColorValue('#json #theme_cc #object #media_col_1', json.theme_cc.object.media_col[1]);
			setColorValue('#json #theme_cc #object #media_col_2', json.theme_cc.object.media_col[2]);
			setColorValue('#json #theme_cc #object #mfilter_col_0', json.theme_cc.object.mfilter_col[0]);
			setColorValue('#json #theme_cc #object #mfilter_col_1', json.theme_cc.object.mfilter_col[1]);
			setColorValue('#json #theme_cc #object #mfilter_col_2', json.theme_cc.object.mfilter_col[2]);
			setColorValue('#json #theme_cc #object #audio_col_0', json.theme_cc.object.audio_col[0]);
			setColorValue('#json #theme_cc #object #audio_col_1', json.theme_cc.object.audio_col[1]);
			setColorValue('#json #theme_cc #object #audio_col_2', json.theme_cc.object.audio_col[2]);
			setColorValue('#json #theme_cc #object #afilter_col_0', json.theme_cc.object.afilter_col[0]);
			setColorValue('#json #theme_cc #object #afilter_col_1', json.theme_cc.object.afilter_col[1]);
			setColorValue('#json #theme_cc #object #afilter_col_2', json.theme_cc.object.afilter_col[2]);
			setColorValue('#json #theme_cc #object #control_col_0', json.theme_cc.object.control_col[0]);
			setColorValue('#json #theme_cc #object #control_col_1', json.theme_cc.object.control_col[1]);
			setColorValue('#json #theme_cc #object #control_col_2', json.theme_cc.object.control_col[2]);
			setColorValue('#json #theme_cc #object #inactive_col_0', json.theme_cc.object.inactive_col[0]);
			setColorValue('#json #theme_cc #object #inactive_col_1', json.theme_cc.object.inactive_col[1]);
			setColorValue('#json #theme_cc #object #inactive_col_2', json.theme_cc.object.inactive_col[2]);
			setColorValue('#json #theme_cc #object #clipping_col', json.theme_cc.object.clipping_col);
			setIntValue('#json #theme_cc #object #clipping_height', json.theme_cc.object.clipping_height);

			if (json.theme_cc.object.midpt_size)
			{
				setIntValue('#json #theme_cc #object #midpt_size_0', json.theme_cc.object.midpt_size[0]);
				setIntValue('#json #theme_cc #object #midpt_size_1', json.theme_cc.object.midpt_size[1]);
				setIntValue('#json #theme_cc #object #midpt_size_2', json.theme_cc.object.midpt_size[2]);
			}

			setColorValue('#json #theme_cc #object #name_col_0', json.theme_cc.object.name_col[0]);
			setColorValue('#json #theme_cc #object #name_col_1', json.theme_cc.object.name_col[1]);
		}

		if (json.theme_cc.timeline)
		{
			// theme_cc timeline

			setColorValue('#json #theme_cc #timeline #scale_col_0', json.theme_cc.timeline.scale_col[0]);
			setColorValue('#json #theme_cc #timeline #scale_col_1', json.theme_cc.timeline.scale_col[1]);
			setColorValue('#json #theme_cc #timeline #bpm_grid_col_0', json.theme_cc.timeline.bpm_grid_col[0]);
			setColorValue('#json #theme_cc #timeline #bpm_grid_col_1', json.theme_cc.timeline.bpm_grid_col[1]);
		}
	}

	if (json.switch) {
		setBoolValue('#json #switch #switch_access_key', json.switch['access_key']);
		setBoolValue('#json #switch #switch_exo_aviutl_filter', json.switch['exo_aviutl_filter']);
		setBoolValue('#json #switch #switch_exo_track_minusval', json.switch['exo_track_minusval']);
		setBoolValue('#json #switch #switch_exo_sceneidx', json.switch['exo_sceneidx']);
		setBoolValue('#json #switch #switch_exo_trackparam', json.switch['exo_trackparam']);
		setBoolValue('#json #switch #switch_exo_specialcolorconv', json.switch['exo_specialcolorconv']);
		setBoolValue('#json #switch #switch_text_op_size', json.switch['text_op_size']);
		setBoolValue('#json #switch #switch_ignore_media_param_reset', json.switch['ignore_media_param_reset']);
		setBoolValue('#json #switch #switch_theme_cc', json.switch['theme_cc']);
		setBoolValue('#json #switch #switch_console', json.switch['console']);
		setBoolValue('#json #switch #switch_console_escape', json.switch['console.escape']);
		setBoolValue('#json #switch #switch_console_input', json.switch['console.input']);
		setBoolValue('#json #switch #switch_console_debug_string', json.switch['console.debug_string']);
		setBoolValue('#json #switch #switch_console_debug_string_time', json.switch['console.debug_string.time']);
		setBoolValue('#json #switch #switch_lua', json.switch['lua']);
		setBoolValue('#json #switch #switch_lua_rand', json.switch['lua.rand']);
		setBoolValue('#json #switch #switch_lua_randex', json.switch['lua.randex']);
		setBoolValue('#json #switch #switch_lua_getvalue', json.switch['lua.getvalue']);
		setBoolValue('#json #switch #switch_lua_path', json.switch['lua.path']);
		setBoolValue('#json #switch #switch_fast', json.switch['fast']);
		setBoolValue('#json #switch #switch_fast_cl', json.switch['fast.cl']);
		setBoolValue('#json #switch #switch_fast_exeditwindow', json.switch['fast.exeditwindow']);
		setIntValue('#json #switch #switch_fast_exeditwindow_step', json.switch['fast.exeditwindow.step']);
		setBoolValue('#json #switch #switch_fast_settingdialog', json.switch['fast.settingdialog']);
		setBoolValue('#json #switch #switch_fast_radiationalblur', json.switch['fast.radiationalblur']);
		setBoolValue('#json #switch #switch_fast_polortransform', json.switch['fast.polortransform']);
	}
}

function saveFile()
{
	var fileName = 'patch.aul.json';

	if (true) // ファイル保存ダイアログを使用する場合は true にする。
	{
		var shell = new ActiveXObject('WScript.Shell');
		var cd = shell.currentDirectory;

		var fso = new ActiveXObject("Scripting.FileSystemObject");
		var path = window.location.pathname.replace('/', '');
		path = fso.GetParentFolderName(path);
		path = '"' + path + '\\file.dialog.exe" -Save ' + '"' + cd + '"';

		var exec = shell.Exec(path);
		fileName = exec.StdOut.ReadLine();

		if (!fileName)
			return;
	}

	// JSON オブジェクトを作成する。
	var json =
	{
		"console" : {
			"visible" : getBoolValue('#json #console #visible'),
			"rect" : [
				getIntValue('#json #console #rect_0'),
				getIntValue('#json #console #rect_1'),
				getIntValue('#json #console #rect_2'),
				getIntValue('#json #console #rect_3')
			],
		},
		"theme_cc" : {
			"layer" : {
				"height_large" : getIntValue('#json #theme_cc #layer #height_large'),
				"height_medium" : getIntValue('#json #theme_cc #layer #height_medium'),
				"height_small" : getIntValue('#json #theme_cc #layer #height_small'),
				"link_col" : getColorValue('#json #theme_cc #layer #link_col'),
				"clipping_col" : getColorValue('#json #theme_cc #layer #clipping_col'),
				"lock_col" : [
					getColorValue('#json #theme_cc #layer #lock_col_0'),
					getColorValue('#json #theme_cc #layer #lock_col_1')
				],
				"hide_alpha" : getFloatValue('#json #theme_cc #layer #hide_alpha')
			},
			"object": {
				"media_col" : [
					getColorValue('#json #theme_cc #object #media_col_0'),
					getColorValue('#json #theme_cc #object #media_col_1'),
					getColorValue('#json #theme_cc #object #media_col_2')
				],
				"mfilter_col" : [
					getColorValue('#json #theme_cc #object #mfilter_col_0'),
					getColorValue('#json #theme_cc #object #mfilter_col_1'),
					getColorValue('#json #theme_cc #object #mfilter_col_2')
				],
				"audio_col" : [
					getColorValue('#json #theme_cc #object #audio_col_0'),
					getColorValue('#json #theme_cc #object #audio_col_1'),
					getColorValue('#json #theme_cc #object #audio_col_2')
				],
				"afilter_col" : [
					getColorValue('#json #theme_cc #object #afilter_col_0'),
					getColorValue('#json #theme_cc #object #afilter_col_1'),
					getColorValue('#json #theme_cc #object #afilter_col_2')
				],
				"control_col" : [
					getColorValue('#json #theme_cc #object #control_col_0'),
					getColorValue('#json #theme_cc #object #control_col_1'),
					getColorValue('#json #theme_cc #object #control_col_2')
				],
				"inactive_col" : [
					getColorValue('#json #theme_cc #object #inactive_col_0'),
					getColorValue('#json #theme_cc #object #inactive_col_1'),
					getColorValue('#json #theme_cc #object #inactive_col_2')
				],
				"clipping_col": getColorValue('#json #theme_cc #object #clipping_col'),
				"clipping_height": getIntValue('#json #theme_cc #object #clipping_height'),
				"midpt_size": [
					getIntValue('#json #theme_cc #object #midpt_size_0'),
					getIntValue('#json #theme_cc #object #midpt_size_1'),
					getIntValue('#json #theme_cc #object #midpt_size_2')
				],
				"name_col" : [
					getColorValue('#json #theme_cc #object #name_col_0'),
					getColorValue('#json #theme_cc #object #name_col_1')
				],
			},
			"timeline": {
				"scale_col" : [
					getColorValue('#json #theme_cc #timeline #scale_col_0'),
					getColorValue('#json #theme_cc #timeline #scale_col_1')
				],
				"bpm_grid_col" : [
					getColorValue('#json #theme_cc #timeline #bpm_grid_col_0'),
					getColorValue('#json #theme_cc #timeline #bpm_grid_col_1')
				]
			}
		},
		"switch" : {
			"access_key" : getBoolValue('#json #switch #switch_access_key'),
			"exo_aviutl_filter" : getBoolValue('#json #switch #switch_exo_aviutl_filter'),
			"exo_track_minusval" : getBoolValue('#json #switch #switch_exo_track_minusval'),
			"exo_sceneidx" : getBoolValue('#json #switch #switch_exo_sceneidx'),
			"exo_trackparam" : getBoolValue('#json #switch #switch_exo_trackparam'),
			"exo_specialcolorconv" : getBoolValue('#json #switch #switch_exo_specialcolorconv'),
			"text_op_size" : getBoolValue('#json #switch #switch_text_op_size'),
			"ignore_media_param_reset" : getBoolValue('#json #switch #switch_ignore_media_param_reset'),
			"theme_cc" : getBoolValue('#json #switch #switch_theme_cc'),
			"console" : getBoolValue('#json #switch #switch_console'),
			"console.escape" : getBoolValue('#json #switch #switch_console_escape'),
			"console.input" : getBoolValue('#json #switch #switch_console_input'),
			"console.debug_string" : getBoolValue('#json #switch #switch_console_debug_string'),
			"console.debug_string.time": getBoolValue('#json #switch #switch_console_debug_string_time'),

			"lua" : getBoolValue('#json #switch #switch_lua'),
			"lua.rand" : getBoolValue('#json #switch #switch_lua_rand'),
			"lua.randex" : getBoolValue('#json #switch #switch_lua_randex'),
			"lua.getvalue": getBoolValue('#json #switch #switch_lua_getvalue'),
			"lua.path" : getBoolValue('#json #switch #switch_lua_path'),

			"fast" : getBoolValue('#json #switch #switch_fast'),
			"fast.cl" : getBoolValue('#json #switch #switch_fast_cl'),
			"fast.exeditwindow" : getBoolValue('#json #switch #switch_fast_exeditwindow'),
			"fast.exeditwindow.step" : getIntValue('#json #switch #switch_fast_exeditwindow_step'),
			"fast.settingdialog" : getBoolValue('#json #switch #switch_fast_settingdialog'),
			"fast.radiationalblur" : getBoolValue('#json #switch #switch_fast_radiationalblur'),
			"fast.polortransform" : getBoolValue('#json #switch #switch_fast_polortransform')
		}
	};

	// JSON オブジェクトを文字列に変換する。
	var from_json = JSON.stringify(json, null , '\t');

	// JSON 文字列をファイルに保存する。
	var fs = new ActiveXObject('Scripting.FileSystemObject');
	var file = fs.OpenTextFile(fileName, 2, true);
	file.Write(from_json);
	file.Close();
}

function refreshPreview()
{
	function getGradient(selector)
	{
		return 'linear-gradient(to right, #' + getColorValue(selector + '_0') + ', #' + getColorValue(selector + '_1') + ')';
	}

	function rgbToHex(rgb)
	{
		function toHex(v) {
			return ('00' + v.toString(16)).substr(-2);
		}
		return '#' + toHex(rgb.r) + toHex(rgb.g) + toHex(rgb.b);
	}

	// タイムラインの高さを設定する。
	var height = Math.round(getIntValue('#json #theme_cc #layer #height_medium'));
	$('.preview div').css('margin-top', '4px');
	$('.preview div').css('font-size', height + 'px');
	$('.preview div').css('line-height', '100%');

	// オブジェクトの文字色を設定する。
	$('.preview th').css('color', '#000');
	$('.preview td').css('color', getPureColorValue('#json #theme_cc #object #name_col_0'));
	$('.preview .inactive-layer th').css('color', getPureColorValue('#json #theme_cc #object #name_col_1'));
	$('.preview .inactive-layer td').css('color', getPureColorValue('#json #theme_cc #object #name_col_1'));

	// オブジェクトの背景色を設定する。
	$('.preview .media-object div').css('background', getGradient('#json #theme_cc #object #media_col'));
	$('.preview .audio-object div').css('background', getGradient('#json #theme_cc #object #audio_col'));
	$('.preview .control-object div').css('background', getGradient('#json #theme_cc #object #control_col'));
	$('.preview .selected-media-object div').css('background-color', getPureColorValue('#json #theme_cc #object #media_col_2'));
	$('.preview .selected-audio-object div').css('background-color', getPureColorValue('#json #theme_cc #object #audio_col_2'));
	$('.preview .selected-control-object div').css('background-color', getPureColorValue('#json #theme_cc #object #control_col_2'));
	$('.preview .mfilter-object div').css('background', getGradient('#json #theme_cc #object #mfilter_col'));
	$('.preview .afilter-object div').css('background', getGradient('#json #theme_cc #object #afilter_col'));
	$('.preview .inactive-object div').css('background', getGradient('#json #theme_cc #object #inactive_col'));
	$('.preview .selected-mfilter-object div').css('background-color', getPureColorValue('#json #theme_cc #object #mfilter_col_2'));
	$('.preview .selected-afilter-object div').css('background-color', getPureColorValue('#json #theme_cc #object #afilter_col_2'));
	$('.preview .selected-inactive-object div').css('background-color', getPureColorValue('#json #theme_cc #object #inactive_col_2'));

	// タイムラインの背景色を設定する。
	var hide_alpha = getFloatValue('#json #theme_cc #layer #hide_alpha');
	var base = { "r":0xf0, "g":0xf0, "b":0xf0 };
	var baseString = rgbToHex(base);
	var inactive = {
		"r":base.r * hide_alpha,
		"g":base.g * hide_alpha,
		"b":base.b * hide_alpha
	};
	var inactiveString = rgbToHex(inactive);
	$('.preview th, .preview td').css('background-color', baseString);
	$('.preview .inactive-layer th, .preview .inactive-layer td').css('background-color', inactiveString);

	// 透明度を設定する。
//	$('.preview .inactive-layer').css('opacity', getFloatValue('#json #theme_cc #layer #hide_alpha'));
}

$(document).ready(function()
{
	// 初期設定。
	$('input[type="text"]').on('input', function(event) {
		refreshPreview();
	});
	$('input[type="color"]').spectrum({
		change: refreshPreview,
		move: refreshPreview,
		preferredFormat: 'hex',
		showInitial: true,
		showInput: true
	});

	// カレントディレクトリを変更する。
	var fso = new ActiveXObject('Scripting.FileSystemObject');
	var shell = new ActiveXObject('WScript.Shell');
	var path = document.URL.replace('file://', '');
	path = fso.GetFile(path).ParentFolder.ParentFolder.Path;
	shell.currentDirectory = path;

	// patch.aul.json ファイルを読み込む。
	loadFile(false);

	// プレビューを更新する。
	refreshPreview();

	// ウィンドウサイズを調整する。
	var contents = document.getElementById('contents');
	var w = contents.scrollWidth;
	var h = contents.scrollHeight;
	var aw = window.outerWidth - window.innerWidth;
	var ah = window.outerHeight - window.innerHeight;
	window.resizeTo(w + aw, h + ah);
});

$('#load').click(function()
{
	// ファイルを読み込む。
	loadFile(true);

	// プレビューを更新する。
	refreshPreview();
});

$('#save').click(function()
{
	// ファイルに保存する。
	saveFile();
});

$('input[type="reset"]').click(function()
{
	// すべての input を初期値に戻す。
	$('#form')[0].reset();

	// プレビューを更新する。
	refreshPreview();
});

});
