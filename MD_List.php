<html>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
<TITLE>製造仕掛不良リスト</TITLE>
<base target="_blank">
<?php	date_default_timezone_set('Asia/Tokyo'); ?>
<script type="text/javascript" src="./MD_List.js?<?php echo date('ymdHi',filemtime('./MD_List.js')); ?>" charset="UTF-8"></script>
<script src="./jquery-4.0.0.min.js"></script>
<script src="./xlsx.full.min.js"></script>
<link id="Style_Sheet"rel="stylesheet"  type="text/css" href="MD_List.css?<?php echo date("ymdHi",filemtime("MD_List.css")); ?>"/>

</HEAD>

<?php
	// エラー表示
	ini_set('display_errors', 1);
	error_reporting(E_ALL);

	// スプレッドシートのID
	$fileId = "1QvHerdYsigDufSJ4d0psbOi2ySms55a29PtJaWU1EPo";

	// スプレッドシートをExcel(xlsx)として書き出すURLの組み立て
	$base  = "https://docs.google.com/spreadsheets/d/";
	$query = "/export" . "?" . "format=xlsx"; // ここでxlsx形式を指定
	$url   = $base . $fileId . $query;

	// ダウンロード実行
	$ch = curl_init();
	curl_setopt($ch, CURLOPT_URL, $url);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
	curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true); // 変換処理のリダイレクトを追う
	curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
	curl_setopt($ch, CURLOPT_USERAGENT, 'Mozilla/5.0');

	$data = curl_exec($ch);
	$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
	curl_close($ch);

	if ($httpCode === 200) {
		$savePath = "downloaded_sheet.xlsx";
		file_put_contents($savePath, $data);
	}
	else {
	}
	
	$filename2 = "/share/external/remotes/unify002-prj/Group/ｾﾝﾀｰ/マウンター関連データー/マウンターライン生産情報/生産情報2020～最新.xlsx";
	$filesave2 = "/share/MD0_DATA/Web/QM/qm-intra/test_PHP_JAVA/SSC_REMAIN/TEST.xlsx";
	
	function my_escapeshellarg($arg) {
		return '"' . str_replace('"', '\\"', $arg) . '"';
	}
	$cmd2 = "cp " . my_escapeshellarg($filename2) . " " . my_escapeshellarg($filesave2);
	system($cmd2);
	chmod($filesave2, 0777);

?>

<body id="body" link=blue vlink=purple onload="FILEopen()">
<div class="no-print">
<h2>製造仕掛不良リスト<p id="L_howto" style="font-size:12px;margin:10px;display:inline;width:fit-content;">使用方法（クリックすると開きます）</p>
</h2>
<div id="tops"  hidden>
<div id="howto" class="howto">
　製造仕掛不良のデータベース（製造にてGoogleドライブ上管理）の簡易閲覧ページ（仮）です。
<ul>
	<li>
		絞り込み検索時は絞り込み検索履歴に追加されます。<p class="del_btn" style="display:inline-block;top:18px"></p>ボタンを押すと対象の検索を除いた検索結果が表示されます。
	</li>
	<li>
		検索対象で「製品名」を選択して検索ワードを空欄で検索すると、全てが検索対象となります。残件一覧の表示に使うと便利かもしれません。
	</li>
	<li>
		URLにハッシュタグ#の後に「検索対象:検索ワード」をつけて検索結果が作成できます。&で繋げば絞り込み検索結果となります。<br>
		例: <a href="http://192.168.130.11/QM/qm-intra/MD_List/MD_List.php#製品名:MV-1200&基板名:MAIN&作番:R24194">http://192.168.130.11/QM/qm-intra/test_PHP_JAVA/MD_List/MD_List.php#製品名:MV-1200&基板名:MAIN&作番:R24194</a>
	</li>
</ul>

</div>
</div>
<div id="F_info">
	<p style="margin:0px 15px" id="F_link"></p>
	<table id="F_TABLE"></table>
</div>
<input type="button" id="all" onclick="fast()" value="直近1週間表示">　※ 直近1週間の受入日で発生した不良情報を表示します。<br>
<h3 style="white-space:nowrap">検索ツール<span class="new_cont">新規/絞り込み検索 <input type="radio" id="new" name="nc" checked onkeypress=enter_key(event) required></input>
			<label for="new">新規</label>
			<input type="radio" id="cont" name="nc" onkeypress=enter_key(event) required></input>
			<label for="cont">絞り込み</label> （全件からの検索、表示中の件の絞り込み検索の選択）</span></h3>
<div style="position:relative">
<div id ="history_box" class="history_box" style="visibility:hidden">
<span style="font-size:12px">絞り込み検索履歴</span>
<table id="History_table"><thead><th></th><th>対象</th><th>表示分</th><th>完全/含む</th><th>複数</th><th>検索ワード</th></thead><tbody></tbody></table>
</div>
</div>


<table id="search">
	<tr>
		<td>
			検索対象
			<select id="SEARCH_TYPE" onkeypress=enter_key(event)>
				<option value="製品名" selected>製品名</option>
				<option value="作番" >作番</option>
				<option value="基板名" >基板名</option>
				<option value="No." >No.</option>
				<option value="シリアルNo." >シリアルNo.</option>
				<option value="受入日" >受入日</option>
				<option value="不良内容/進捗状況" >不良内容/進捗状況</option>
				<option value="図番" >図番</option>
				<option value="不良IC 不良部品/部品" >不良IC 不良部品/部品</option>
			</select>
		</td>
		<td>
			<input id="SEARCH_TXT" type="text" onkeypress=enter_key(event) "focus"></input>
			<input type="button" id="SEARCH_GO"  onclick= SEARCH_ON("NEW") value="検索"></input>
		</td>
		<td><span id="exam1">（スペース区切りで複数ワードでの検索可能）</span></td>
	</tr>

	<tr>
		<td style="text-align:right;">残件/完了表示</td>
		<td>
			<input type="checkbox" id="zan" name="rem" checked onkeypress=enter_key(event)></input>
			<label for="zan">残件分</label>
			<input type="checkbox" id="complete" name="rem" checked onkeypress=enter_key(event)></input>
			<label for="complete">完了分</label>
		</td>
		<td><span id="zan_comp">（残件分、完了分の表示/非表示の選択）</span></td>
	</tr>
	<tr>
		<td style="text-align:right;">入力したワード</td>
		<td>
			<input type="radio" id="kan" name="aaa" value="完全一致" onkeypress=enter_key(event)></input>
			<label for="kan">完全一致</label>
			<input type="radio" id="bubun" name="aaa" value="含む" checked onkeypress=enter_key(event)></input>
			<label for="bubun">含む</label>
			<input type="radio" id="non_bubun" name="aaa" value="含まない" onkeypress=enter_key(event)></input>
			<label for="non_bubun">含まない</label>
		</td>
		<td><span id="exam2">（ワードと完全一致/ワードを含む項目を検索）</span></td>
	</tr>
	<tr>
		<td style="text-align:right;">複数ワード</td>
		<td>
			<input type="radio" id="or" name="bbb" value="いずれか" checked onkeypress=enter_key(event)></input>
			<label for="or" value="OR">いずれか</label>
			<input type="radio" id="and" name="bbb" value="全て" onkeypress=enter_key(event)></input>
			<label for="and" value="AND">全て</label>
		</td>
		<td><span id="exam3">（複数ワードでの検索時、いずれか/全てに該当する項目を検索）</span></td>
	</tr>
</table>
<span id="TITLE1" style="margin: 10px 0;visibility: hidden;">
    読み込み中...: <span id="progress-percent">0</span>% 
</span>
</div>
<div id="TITLE2" style="margin: 0px 0px 0px 0px; visibility: hidden;" >　</div>
<table id="R_LIST"></table>

<div id="toTop" style="display:none" class="no-print"><div id="delta"></div><div id="sq"></div></div>
<!--<div id="test">test</div>-->
</body>
<!--Version 1.1 .js、xlsxのキャッシュ対策のため.phpへ変更 -->
<!--Version 1.2 .js、xlsxのURLのアンカーの説明追加 -->
<!--Version 1.3 アップロードページへのリンク追加 -->
<!--Version 1.4 残件＆社外内訳＆完了日数＆総滞留日数グラフページへのリンク追加 -->
<!--Version 1.5 表示中残件数を追加 -->
<!--Version 1.6 共有サーバー（\\192.168.130.17（Unify002）上のファイルをWEBサーバー上にコピーして使用するよう修正。共有サーバーへのアクセスはVS山下さんにサーバー側の設定をして頂いた。今後、サーバー更新等で設定が引き継がれない可能性もあるため注意 アップロードページへのリンクも削除。エクセルファイルがおかしいと「読み込み中…」から「正しい形式のエクセルファイルが保存されていません。」と表示が変わるよう変更-->
<!--Version 2.0 ファイル構成の変更につき、script typeにファイルを追加。ハッシュタグ検索で#?製品番号対応したことを追記-->
<!--Version 2.1 PHPのcopy関数使用を止めてsystem（Linux）のcpコマンドで動作するようにした。これでファイルコピー時の謎のファイル破損が起こらないと思われる。どうもエクセルファイルの保存時に内部不整合が発生したりするらしい。EXCELアプリはその辺補完するけど、PHPにはできず結果ファイル破損が起こる可能性があるらしい-->
</html>
