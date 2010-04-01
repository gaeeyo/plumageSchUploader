// PLUMAGE の予約リストをしょぼいカレンダーにアップロードするツール
//
//  cscript plumageSchUploader.js <user> <pass> [slot]
//
//   <user>  しょぼいカレンダーのUserID
//   <pass>  しょぼいカレンダーのパスワード
//   [slot]  0～3の数値(未指定の場合0)
// 
// PLUMAGE.dat があるディレクトリで実行してください。
// 7日前から3週間分をアップロードします。
// 

var _userAgent = 'plumageSchUploader/1.0.0';
var _uploadUrl = 'http://cal.syoboi.jp/sch_upload';
var _trace = 0;	// デバッグ出力
var _day   = 24*60*60*1000;		// 1日の時間

// 好みで変えて良し
var _scope_start = -7 * _day;	// 何日前から出力するか(-7*_day で 7日前から出力)
var _scope_span  = 22 * _day;	// 何日間出力するか(15*_day で15日分出力)
var _devName = 'PL';					// デバイス名(しょぼいカレンダー上に表示される)

main(WScript.Arguments);

function main(args)
{
	if (args.length < 2) {
		WScript.Echo('plumageSchUploader.js <user> <pass> [slot]');
	}
	else {
		var conf = loadPlumageConf('PLUMAGE.conf');
		var items = loadSchFile('PLUMAGE.dat', conf);
		var sch_data = formatItems(items);
		var slot = (args.length > 2 ? args(2) : 0);
		
		upload(args(0), args(1), sch_data, slot);
	}
}

function XmlUtil(node)
{
	this.node = node;
	
	this.attach = function (node) {
		this.node = node;
	}
	
	this.text2bool = function(text) {
		return (text.match(/^(0|false)$/i) ? false : true);
	}
	
	this.text2date = function(text) {
		var m = text.match(/^(\d+)-(\d+)-(\d+)T(\d+):(\d+):(\d\d)/);
		if (!m) WScript.Echo('text2date error : '+text);
		return new Date(m[1], m[2]-1, m[3], m[4], m[5], m[6]);
	}

	this.getTime = function(exp) {
		var text = this.getText(exp);
		var m = text.match(/^\d+-\d+-\d+T(\d+):(\d+):(\d+)/);
		return m[1]*60*60*1000 + m[2]*60*1000 + m[3]*1000;
	}
	
	this.getDate = function(exp) {
		return this.text2date(this.getText(exp));
	}
	
	this.getBool = function(exp) {
		return this.text2bool(this.getText(exp));
	}
	
	this.getText = function(exp) {
		var n = this.node.selectSingleNode(exp);
		if (n) {
			return n.text;
		}
		return '';
	}
	
	this.getBoolArray = function(exp) {
		var ns = this.node.selectNodes(exp);
		var a = [];
		for (var j=0; j<ns.length; j++) {
			a.push(this.text2bool(ns(j).text));
		}
		return a;
	}
}

function trace()
{
	if (_trace) {
		var text = [];
		for (var j=0; j<arguments.length; j++) {
			text.push(arguments[j]);
		}
		WScript.Echo(text.join(' '));
	}
}

// PLUMAGE.conf の読み込み
function loadPlumageConf(path)
{
	var chMap = {};
	
	path = path.replace('.dat', '.conf');
	var xml = new ActiveXObject('MSXML2.DOMDocument');
	if (!xml.load(path)) {
		WScript.Echo('LOAD ERROR: '+path);
		WScript.Quit(1);
	}

	var chConvs = xml.selectNodes('/ClsProp/chConv/CHCONV');
	var x = new XmlUtil();
	for (var j=0; j<chConvs.length; j++) {
		x.attach(chConvs(j));
		chMap[x.getText('Remote')] = x.getText('Station');
		trace(x.getText('Remote') + ': ' + x.getText('Station'));
	}
	
	return {
		ch: chMap
	};
}

// PLUMAGE.dat の読み込み
function loadSchFile(path, conf)
{
	// 出力範囲を計算、7日前の0:00から7日後まで
	var scope_start = new Date((new Date()).getTime() + _scope_start);
	scope_start = new Date(scope_start.getFullYear(), scope_start.getMonth(),
		scope_start.getDate());
	var scope_end = new Date(scope_start.getTime() + _scope_span);
	
	trace('scope:', dateStr(scope_start), '-', dateStr(scope_end));

	// plumage.dat を読む込む
	var xml = new ActiveXObject('MSXML2.DOMDocument');
	if (!xml.load(path)) {
		WScript.Echo('LOAD ERROR: '+path);
		WScript.Quit(1);
	}
	
	var items = [];
	var tts = xml.selectNodes('/ClsTimeList/tt/TimeTable');
	var x = new XmlUtil();
	for (var j=0; j<tts.length; j++) {
		x.attach(tts(j));
		
		if (x.getBool('isDisable')) continue; // [予約無効] を無視
		
		var nextDay = x.getDate('NextDay');
		var program = x.getText('Program');
		var bgnTime = x.getTime('BgnTime');
		var endTime = x.getTime('EndTime');
		var spnTime = (endTime > bgnTime ? endTime : endTime + 24*60*60*1000) - bgnTime;
		var input = x.getText('Input');
		var remote = x.getText('Remote');
		
		var station = (typeof(conf.ch[remote]) != 'undefined' ? conf.ch[remote] : 'unknown');
		//   mapStation(getText(t,'Input'), getText(t,'Remote'));
		
		if (x.getBool('isWeekly')) {
			// 毎週の予定
			if (nextDay >= scope_end) {
				trace('まだ先:'+program);
				continue;	// 範囲外のものを無視
			}
			
			// 終了日の指定がある場合に確認
			var endDate = x.getDate('EndDate');
			var useEndDate = x.getBool('useEndDate');
			if (useEndDate && endDate < scope_start) {
				trace('終わってる:'+program);
				continue; // 終わってるもの無視
			}

			// [true,false,...] みたいな曜日毎のチェック
			var wlist = x.getBoolArray('WList/boolean');
			for (var d = new Date(scope_start); d < scope_end; d.setTime(d.getTime() + _day)) {
				if (nextDay > d || !wlist[d.getDay()]) {
					continue;
				}
				var start = new Date(d.getTime() + bgnTime);
				var item = {
					START: start.getTime()/1000,
					END: (start.getTime() + spnTime)/1000,
					TITLE: program,
					SUBTITLE: '',
					DEV: _devName,
					STATION: station
				};
				dumpItem(item);
				items.push(item);
			}
		}
		else {
			// 1回きりの予定
			var start = new Date(nextDay.getTime() + bgnTime);
			var end = new Date(start.getTime() + spnTime);
			var item = {
				START: start.getTime() / 1000,
				END: end.getTime() / 1000,
				TITLE: program,
				SUBTITLE: '',
				DEVNO: _devName,
				STATION: station
			};
			dumpItem(item);
			items.push(item);
		}
	}
	return items;
	
	function dumpItem(item) {
		var st = new Date(item.START * 1000);
		trace(dateStr(st),
			(item.END - item.START)/60 + '分',
			item.TITLE,
			item.STATION
		);
	}
}

function dateStr(d) {
	function fn(a) {
		return ('0'+a).slice(-2);
	}
	return	[d.getFullYear(),fn(d.getMonth()+1),fn(d.getDate())].join('-')
		+'('
		+('日月火水木金土'.substr(d.getDay(),1))
		+') '
		+[fn(d.getHours()), fn(d.getMinutes()), fn(d.getSeconds())].join(':');
}

// アップロードするデータの形式に変換(csvに)
function formatItems(items)
{
	function tsvEscape(text) {
		return text.replace("\t", ' ');
	}

	var text = '';
	for (var j in items) {
		var item = items[j];
		text += [
			item.START,
			item.END,
			item.DEV,
			tsvEscape(item.TITLE),
			tsvEscape(item.STATION),
			tsvEscape(item.SUBTITLE),
			0,
			0
		].join("\t")+"\n";
	}
	//WScript.Echo(text); WScript.Quit(0);
	return text;
	
}

// アップロード
function upload(user, pass, sch_data, slot)
{
	var http = new ActiveXObject('MSXML2.XMLHTTP');
	
	http.Open('POST', _uploadUrl+'?slot='+slot, false, user, pass);
	http.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
	http.setRequestHeader('User-agent', _userAgent);
	http.onreadystatechange = function(){
		if (http.readyState == 4) {
			if (http.status == 200) {
				WScript.Echo(http.responseText);
			}
			else {
				WScript.Echo('UPLOAD ERROR: '+http.status);
			}
		}
	};
	http.send(
		'data='+encodeURIComponent(sch_data)
	);
}
