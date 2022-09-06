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
		alert(fileName + ' を開けませんでした');
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

	if (json.redo) {
		setBoolValue('#redo_shift', json.redo['shift']);
	}

	if (json.fast_exeditwindow) {
		setIntValue('#fast_exeditwindow_step', json.fast_exeditwindow['step']);
	}

	if (json.fast_text) {
		setIntValue('#fast_text_release_time', json.fast_text['release_time']);
	}

	if (json.switch) {
		setBoolValue('#switch_access_key', json.switch['access_key']);
		setBoolValue('#switch_exo_aviutl_filter', json.switch['exo_aviutl_filter']);
		setBoolValue('#switch_exo_sceneidx', json.switch['exo_sceneidx']);
		setBoolValue('#switch_exo_trackparam', json.switch['exo_trackparam']);
		setBoolValue('#switch_exo_track_minusval', json.switch['exo_track_minusval']);
		setBoolValue('#switch_exo_specialcolorconv', json.switch['exo_specialcolorconv']);
		setBoolValue('#switch_tra_aviutlfilter', json.switch['tra_aviutlfilter']);
		setBoolValue('#switch_tra_change_drawfilter', json.switch['tra_change_drawfilter']);
		setBoolValue('#switch_tra_specified_speed', json.switch['tra_specified_speed']);
		setBoolValue('#switch_text_op_size', json.switch['text_op_size']);
		setBoolValue('#switch_ignore_media_param_reset', json.switch['ignore_media_param_reset']);
		setBoolValue('#switch_failed_sjis_msgbox', json.switch['failed_sjis_msgbox']);
		setBoolValue('#switch_theme_cc', json.switch['theme_cc']);
		setBoolValue('#switch_exeditwindow_sizing', json.switch['exeditwindow_sizing']);
		setBoolValue('#switch_settingdialog_move', json.switch['settingdialog_move']);
		setBoolValue('#switch_obj_colorcorrection', json.switch['obj_colorcorrection']);
		setBoolValue('#switch_obj_lensblur', json.switch['obj_lensblur']);
		setBoolValue('#switch_obj_noise', json.switch['obj_noise']);
		setBoolValue('#switch_settingdialog_excolorconfig', json.switch['settingdialog_excolorconfig']);
		setBoolValue('#switch_r_click_menu_split', json.switch['r_click_menu_split']);
		setBoolValue('#switch_r_click_menu_delete', json.switch['r_click_menu_delete']);
		setBoolValue('#switch_blend', json.switch['blend']);
		setBoolValue('#switch_undo', json.switch['undo']);
		setBoolValue('#switch_undo_redo', json.switch['undo.redo']);
		setBoolValue('#switch_console', json.switch['console']);
		setBoolValue('#switch_console_escape', json.switch['console.escape']);
		setBoolValue('#switch_console_input', json.switch['console.input']);
		setBoolValue('#switch_console_debug_string', json.switch['console.debug_string']);
		setBoolValue('#switch_console_debug_string_time', json.switch['console.debug_string.time']);
		setBoolValue('#switch_lua', json.switch['lua']);
		setBoolValue('#switch_lua_env', json.switch['lua.env']);
		setBoolValue('#switch_lua_path', json.switch['lua.path']);
		setBoolValue('#switch_lua_getvalue', json.switch['lua.getvalue']);
		setBoolValue('#switch_lua_rand', json.switch['lua.rand']);
		setBoolValue('#switch_lua_randex', json.switch['lua.randex']);
		setBoolValue('#switch_fast', json.switch['fast']);
		setBoolValue('#switch_fast_exeditwindow', json.switch['fast.exeditwindow']);
		setBoolValue('#switch_fast_settingdialog', json.switch['fast.settingdialog']);
		setBoolValue('#switch_fast_text', json.switch['fast.text']);
		setBoolValue('#switch_fast_create_figure', json.switch['fast.create_figure']);
		setBoolValue('#switch_fast_border', json.switch['fast.border']);
		setBoolValue('#switch_fast_cl', json.switch['fast.cl']);
		setBoolValue('#switch_fast_radiationalblur', json.switch['fast.radiationalblur']);
		setBoolValue('#switch_fast_polortransform', json.switch['fast.polortransform']);
		setBoolValue('#switch_fast_displacementmap', json.switch['fast.displacementmap']);
		setBoolValue('#switch_fast_flash', json.switch['fast.flash']);
		setBoolValue('#switch_fast_directionalblur', json.switch['fast.directionalblur']);
		setBoolValue('#switch_fast_lensblur', json.switch['fast.lensblur']);
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
		"redo" : {
			"shift" : getBoolValue('#redo_shift')
		},
		"fast_exeditwindow" : {
			"step" : getIntValue('#fast_exeditwindow_step')
		},
		"fast_text" : {
			"release_time" : getIntValue('#fast_text_release_time')
		},
		"switch" : {
			"access_key" : getBoolValue('#switch_access_key'),
			"exo_aviutl_filter" : getBoolValue('#switch_exo_aviutl_filter'),
			"exo_sceneidx" : getBoolValue('#switch_exo_sceneidx'),
			"exo_trackparam" : getBoolValue('#switch_exo_trackparam'),
			"exo_track_minusval" : getBoolValue('#switch_exo_track_minusval'),
			"exo_specialcolorconv" : getBoolValue('#switch_exo_specialcolorconv'),
			"tra_aviutlfilter" : getBoolValue('#switch_tra_aviutlfilter'),
			"tra_change_drawfilter": getBoolValue('#switch_tra_change_drawfilter'),
			"tra_specified_speed": getBoolValue('#switch_tra_specified_speed'),
			"text_op_size" : getBoolValue('#switch_text_op_size'),
			"ignore_media_param_reset" : getBoolValue('#switch_ignore_media_param_reset'),
			"failed_sjis_msgbox": getBoolValue('#switch_failed_sjis_msgbox'),
			"theme_cc" : getBoolValue('#switch_theme_cc'),
			"exeditwindow_sizing" : getBoolValue('#switch_exeditwindow_sizing'),
			"settingdialog_move" : getBoolValue('#switch_settingdialog_move'),
			"obj_colorcorrection": getBoolValue('#switch_obj_colorcorrection'),
			"obj_lensblur": getBoolValue('#switch_obj_lensblur'),
			"obj_noise": getBoolValue('#switch_obj_noise'),
			"settingdialog_excolorconfig": getBoolValue('#switch_settingdialog_excolorconfig'),
			"r_click_menu_split": getBoolValue('#switch_r_click_menu_split'),
			"r_click_menu_delete": getBoolValue('#switch_r_click_menu_delete'),
			"blend": getBoolValue('#switch_blend'),
			"undo": getBoolValue('#switch_undo'),
			"undo.redo": getBoolValue('#switch_undo_redo'),
			"console" : getBoolValue('#switch_console'),
			"console.escape" : getBoolValue('#switch_console_escape'),
			"console.input" : getBoolValue('#switch_console_input'),
			"console.debug_string" : getBoolValue('#switch_console_debug_string'),
			"console.debug_string.time": getBoolValue('#switch_console_debug_string_time'),
			"lua" : getBoolValue('#switch_lua'),
			"lua.env" : getBoolValue('#switch_lua_env'),
			"lua.path" : getBoolValue('#switch_lua_path'),
			"lua.getvalue": getBoolValue('#switch_lua_getvalue'),
			"lua.rand" : getBoolValue('#switch_lua_rand'),
			"lua.randex" : getBoolValue('#switch_lua_randex'),
			"fast" : getBoolValue('#switch_fast'),
			"fast.exeditwindow" : getBoolValue('#switch_fast_exeditwindow'),
			"fast.settingdialog" : getBoolValue('#switch_fast_settingdialog'),
			"fast.text" : getBoolValue('#switch_fast_text'),
			"fast.create_figure": getBoolValue('#switch_fast_create_figure'),
			"fast.border" : getBoolValue('#switch_fast_border'),
			"fast.cl" : getBoolValue('#switch_fast_cl'),
			"fast.radiationalblur" : getBoolValue('#switch_fast_radiationalblur'),
			"fast.polortransform" : getBoolValue('#switch_fast_polortransform'),
			"fast.displacementmap": getBoolValue('#switch_fast_displacementmap'),
			"fast.flash" : getBoolValue('#switch_fast_flash'),
			"fast.directionalblur" : getBoolValue('#switch_fast_directionalblur'),
			"fast.lensblur" : getBoolValue('#switch_fast_lensblur')
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

	$('.preview > div > div').css('font-size', height - 4 + 'px');

	$('.preview > div > div > span').css('margin-top', '4px');
	$('.preview > div > div > span').css('font-size', height - 4 + 'px');
//	$('.preview > div > div > span').css('line-height', '100%');

	// オブジェクトの文字色を設定する。
	$('.preview .active-layer').css('color', getPureColorValue('#json #theme_cc #object #name_col_0'));
	$('.preview .active-layer .layer-title').css('color', '#000');
	$('.preview .inactive-layer').css('color', getPureColorValue('#json #theme_cc #object #name_col_1'));
	$('.preview .inactive-layer').css('color', getPureColorValue('#json #theme_cc #object #name_col_1'));

	// オブジェクトの背景色を設定する。
	$('.preview .media-object span').css('background', getGradient('#json #theme_cc #object #media_col'));
	$('.preview .audio-object span').css('background', getGradient('#json #theme_cc #object #audio_col'));
	$('.preview .control-object span').css('background', getGradient('#json #theme_cc #object #control_col'));
	$('.preview .selected-media-object span').css('background-color', getPureColorValue('#json #theme_cc #object #media_col_2'));
	$('.preview .selected-audio-object span').css('background-color', getPureColorValue('#json #theme_cc #object #audio_col_2'));
	$('.preview .selected-control-object span').css('background-color', getPureColorValue('#json #theme_cc #object #control_col_2'));
	$('.preview .mfilter-object span').css('background', getGradient('#json #theme_cc #object #mfilter_col'));
	$('.preview .afilter-object span').css('background', getGradient('#json #theme_cc #object #afilter_col'));
	$('.preview .inactive-object span').css('background', getGradient('#json #theme_cc #object #inactive_col'));
	$('.preview .selected-mfilter-object span').css('background-color', getPureColorValue('#json #theme_cc #object #mfilter_col_2'));
	$('.preview .selected-afilter-object span').css('background-color', getPureColorValue('#json #theme_cc #object #afilter_col_2'));
	$('.preview .selected-inactive-object span').css('background-color', getPureColorValue('#json #theme_cc #object #inactive_col_2'));

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
	$('.preview .active-layer').css('background-color', baseString);
	$('.preview .inactive-layer').css('background-color', inactiveString);

	// 透明度を設定する。
//	$('.preview .inactive-layer').css('opacity', getFloatValue('#json #theme_cc #layer #hide_alpha'));
}

$(function(){

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
/*
	// ウィンドウサイズを調整する。
	var contents = document.getElementById('contents');
	var w = contents.scrollWidth;
	var h = contents.scrollHeight;
	var aw = window.outerWidth - window.innerWidth;
	var ah = window.outerHeight - window.innerHeight;
	window.resizeTo(w + aw, h + ah);
*/
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

	// spectrum をリセットする。
	$('input[type="color"]').spectrum('destroy');
	$('input[type="color"]').spectrum({
		change: refreshPreview,
		move: refreshPreview,
		preferredFormat: 'hex',
		showInitial: true,
		showInput: true
	});

	// プレビューを更新する。
	refreshPreview();
});

});
