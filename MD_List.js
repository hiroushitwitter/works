//Version 1.0 新規。

const $Id = (id) => document.getElementById(id);
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

let FList ={};
let FList_out = [];
let Disp_count = 0;
let Search_History = [];
let Header ;
let dataBody
let Result;
let prev_search_type;
let test_header;
let test_out_group2;
async function FILEopen(filename){
	FuncloadEventListener();
	
	const testa = await fetch('./downloaded_sheet.xlsx',{ cache: 'no-cache' });
	const testb = await testa.arrayBuffer();
	const EXCELSHEET = XLSX.read(testb, {type: 'array'});

	const RAW_DATA_remain =  EXCELSHEET.Sheets[EXCELSHEET.SheetNames[0]];
	const RAW_DATA_complete =  EXCELSHEET.Sheets[EXCELSHEET.SheetNames[1]];
	const RAW_DATA_old =  EXCELSHEET.Sheets[EXCELSHEET.SheetNames[3]];
	
	let dataBody_remain = XLSX.utils.sheet_to_json(RAW_DATA_remain, { header: 1, defval: "", raw: true, range: 1 });
	let dataBody_complete = XLSX.utils.sheet_to_json(RAW_DATA_complete, { header: 1, defval: "", raw: true });
	let dataBody_old = XLSX.utils.sheet_to_json(RAW_DATA_old, { header: 1, defval: "", raw: true });
	dataBody_old.forEach(a => {
		a.splice(3,0,"");
		a.splice(10,1);
		a.push("","完了");
	});
	dataBody_remain.forEach(a => a.push("残件"));
	dataBody_complete.forEach(a => a.push("完了"));
	dataBody = dataBody_remain.concat(dataBody_complete.slice(1)).concat(dataBody_old.slice(1)).filter(a => a[0]);
	dataBody.forEach((a,i) => a.splice(-1,0,i));
	Header = dataBody[0].concat();
	Making_table(dataBody);

	const FILE_DATA2 = await fetch("TEST.xlsx").then(response => response.arrayBuffer());
	const EXCELSHEET2 = XLSX.read(FILE_DATA2, {type: 'array'});
	const test_DATA = EXCELSHEET2.Sheets[EXCELSHEET2.SheetNames[0]];
	const test_dataBody = XLSX.utils.sheet_to_json(test_DATA, { header: 1, defval: ""}).slice(5);
	test_header = test_dataBody[0].splice(0,24);
	const test_out = test_dataBody.filter(a => {
		if(a[0] && /^\d+$/.test(a[0])){
			a.splice(24);
			return true;
		}
		return false;
		});
	const test_out_sort = test_out.sort((a,b) =>a[5].localeCompare(b[5]));
	const test_out_group =  Object.groupBy(test_out, (a) => a[5]);
	Object.entries(test_out_group).forEach(a => {
		a[1].sort((b,c) => b[8].localeCompare(c[8]));
	});
	console.log("change1")
	test_out_group2 = Object.fromEntries(
		Object.entries(test_out_group).map(([key, value]) => [key,Object.groupBy(value, (b) => b[8].replace(/[^\x00-\x7E]/g, '').trim();)])
	);
}

async function Making_table(data){
	document.querySelectorAll("input,select").forEach(a => a.disabled = 1);

	const FuncSheetLoad = (Indata) => {
		L_data = {};
		Indata.forEach(a => {
			if (!a.some(c => c.length)) return;
			const name = a[4];
			const CD = a[5];
			const R_No = a[2];
			((L_data[name] ??= {})[R_No] ??= {})[CD] ??= [];
			L_data[name][R_No][CD].push(a);	//データ構造としては製品名-R作番-基板名-以下は生データ（製品名、R作番、基板名含む）+通し番号+残件/完了となる。
		});
		return L_data;
	}
	const FList = FuncSheetLoad(data.slice(1));

	//fetch("save.php?name=MD_List_OUT",{
	//	method: "POST",
	//	body: JSON.stringify(FList)
	//});
		
	FList_out = Object.entries(FList).map(a => [a[0],Object.entries(a[1]).map(b => [b[0],Object.entries(b[1]).map(c => [c[0],c[1]])])]);
	Header.splice(0,6,Header[4],Header[2],Header[5],"No.",Header[1],Header[3]);
	Header = Header.filter((a,i) => !(i == 5 || i == 7 || (i>=13 && i<=26)))
	$Id("R_LIST").innerHTML =`<thead><tr><th>${Header.slice(0,-2).join("</th><th>")}</th></tr></thead><tbody></tbody>`;

	await new Promise(resolve => requestAnimationFrame(resolve));
	
	document.querySelectorAll("input,select").forEach(a => a.disabled = 0);
	hash_tag_filter();
	$Id("SEARCH_TXT").focus();
}

async function SEARCH_ON(add_del){

	$Id("new").checked ? Search_History = [] : "";
	Result = FList_out;
	let Table_Buf ="";
	const Input_set = [$Id("SEARCH_TYPE").value,$Id("SEARCH_TXT").value,document.querySelector("input[name='bbb']:checked").value,document.querySelector("input[name='aaa']:checked").value,$Id("zan").checked,$Id("complete").checked];
	if(add_del == "NEW"){
		if(Search_History.length == 0 || !Search_History.some(a => a.join() == Input_set.join())){
			Search_History.push(Input_set);
		}
		else{
			return;
		}
	}
	else if(add_del != "HASH"){
		Search_History.splice((add_del.closest("tr").rowIndex-1),1);
	}
	Search_History.forEach(history => {
		const Input = history[1].replace(/[！-～]/g, (a) => String.fromCharCode(a.charCodeAt(0) - 0xFEE0)).replace(/　/g, " ");
		const Numbers = Input.replace(/[^0-9 ]/g,"").split(/ +/).filter(Boolean);
		const texts = Input.toLowerCase().split(" ").filter(Boolean);
		const textsfull = history[1].replace(/　/g, " ").split(" ").filter(Boolean);
		const excelDate = Math.floor(Date.parse(Input) / 86400000) + 25569 + 1;
		const In_date = new Date(Input);
		const word_checks = (items,name) => items[history[2]  == "いずれか" ? "some" : "every"](item => {
			if(history[3] == "完全一致"){
				return name == item ;
			}
			else{
				return history[3] == "含む" ? name.includes(item) : !name.includes(item);
			}
		});
		const S_col = {"No.":[textsfull,0],"シリアルNo.":[texts,6],"不良内容/進捗状況":[textsfull,[8,9]],"図番":[textsfull,10],"不良IC 不良部品/部品":[textsfull,[11,12]]};
	
		const remain_end = [history[4] ? "残件" : "",history[5] ? "完了" :""];
		Result = Result.flatMap(a => {
			let Result_Product = a[1].flatMap(b => {
				let Result_Rot = b[1].flatMap(c => {
					let Result_CD = c[1].filter(d => remain_end.includes(d[d.length-1]));
					if(Object.keys(S_col).includes(history[0])){
						if(typeof(S_col[history[0]][1]) == "number"){
							Result_CD = Result_CD.filter(d => word_checks(S_col[history[0]][0],String(d[S_col[history[0]][1]])))
						}
						else{
							Result_CD = Result_CD.filter(d => S_col[history[0]][1].some(el => word_checks(S_col[history[0]][0],String(d[el]))));
						}
					}
					else if(history[0] == "受入日"){
						Result_CD = history[3] == "完全一致" ? Result_CD.filter(d => excelDate > d[1]) : Result_CD.filter(d => excelDate < d[1]);
					}
					return  Result_CD.length ? [[c[0],Result_CD]] : [];
				});
				history[0] == "基板名" ? Result_Rot = Result_Rot.filter(c => word_checks(texts,c[0].toLowerCase())) : "";
				return  Result_Rot.length ? [[b[0],Result_Rot]] : [];
			});
			history[0] == "作番" ? Result_Product = Result_Product.filter(b => word_checks(Numbers,b[0])) : "";
			return  Result_Product.length ? [[a[0],Result_Product]] : [];
		});
		history[0] == "製品名" ? Result = texts.length ? Result.filter(a => word_checks(texts,a[0].toLowerCase())) : Result : "";

		Table_Buf += `
			<tr>
				<td><span onclick="SEARCH_ON(this)")><p class="del_btn"></p></span></td>
				<td>${history[0]}</td>
				<td>${history[4] ? "残件": "－"}/${history[5] ?  "完了": "－"}</td>
				<td>${history[0] != "受入日" ? history[3] : (history[3] == "完全一致" ?"以前":"以降")}</td>
				<td>${history[0] != "受入日" ? history[2] : "-"}</td>
				<td>${history[1]}</td>
			</tr>`;
	});
	$Id("History_table").tBodies[0].innerHTML = Table_Buf;
	$Id("History_table").tBodies[0].querySelectorAll("td:last-child").forEach(a => {
		const limit = a.clientWidth;
		const chars = a.textContent;
		if(a.scrollWidth > a.clientWidth){
			for(i=1;i<chars.length;i++){
				a.textContent = chars.slice(0,i)+"...";
				if(a.scrollWidth > limit){
					a.innerHTML = "<span title='"+chars+"'>"+a.textContent.slice(0,i-1)+"...</span>";
					return;
				}
			}
		}
	});
	
	const TableResult = Result.flatMap(a => {
		let Result_Product = a[1].flatMap(b => {
			let Result_Rot = b[1].flatMap(c => {
				let Result_CD = c[1].map(d => d.filter((_, i) => !((i >= 2 && i <= 5) || i == 7 || (i >= 13 && i <= 26))));
				return  Result_CD.length ? [[c[0],Result_CD]] : [];
			});
			return  Result_Rot.length ? [[b[0],Result_Rot]] : [];
		});
		return  Result_Product.length ? [[a[0],Result_Product]] : [];
	});

	Disp_count = 0;
	await new Promise((resolve) => {
		let currentRow = 0;
		renderNextChunk(TableResult,currentRow,resolve);
	});

	$Id("TITLE1").style.visibility = "visible";
	const cur_disable = document.querySelectorAll("input:not(:disabled),select");
	cur_disable.forEach(a => a.disabled = 1);
	requestAnimationFrame(() => setTimeout(() =>{
		$Id("TITLE1").style.visibility = "hidden";
		cur_disable.forEach(a => a.disabled = 0);
		$Id("SEARCH_TXT").focus();
		$Id("TITLE2").style.visibility = "visible";
		$Id("TITLE2").innerHTML = Disp_count == 0 ? `<div style="margin:0px 5px 0px 20px">検索結果：${Disp_count}件</div>` : `<div style="margin:0px 5px 0px 20px">検索結果：${Disp_count}件<button  onclick="FuncDL()">Download</button>※検索結果の詳細をエクセルでダウンロード</div>`;
	}, 0));

}

async function hash_tag_filter(){
	// --- ハッシュタグ検索機能。&でつないだ場合は最後のもので検索（検索ボックス1つだけのため） ---
	const hash_opt = Array.from($Id("SEARCH_TYPE").options).map(opt => opt.value);
	if(location.hash){
		$Id("cont").checked = 1;
		const In_hash_arr = decodeURIComponent(location.hash.slice(1)).split("&");
		for(const In_hash of In_hash_arr){
			let [h_id,h_char] = In_hash.split(":");
			if(hash_opt.includes(h_id)){
				Search_History.push([h_id,h_char,document.querySelector("input[name='bbb']:checked").value,document.querySelector("input[name='aaa']:checked").value,$Id("zan").checked,$Id("complete").checked]);
			}
		}
		SEARCH_ON("HASH");
		$Id("history_box").style.visibility = "visible";
	}
}
function FuncDL(){
	const FileResult = Result.flatMap(a => a[1].flatMap(b =>b[1].flatMap(c => c[1].map(d => {
		d.splice(-2,1);
		return d;
	}))));
	let data = [];
	data.push(dataBody[0].slice(0,-2));
	data = data.concat(FileResult);
	const ws = XLSX.utils.json_to_sheet(data, { skipHeader: true });
	const range = XLSX.utils.decode_range(ws['!ref']);
	for (let i = range.s.r + 1; i <= range.e.r; i++) {
		[1,15,17,19,21,23,25].forEach(a => {
			const cellAddress = XLSX.utils.encode_cell({ r: i, c: a });
			ws[cellAddress] ? ws[cellAddress].z = 'yyyy/mm/dd' : "";
		});
    }
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
	XLSX.writeFile(wb, "export.xlsx"); // ファイルの保存
}
	
async function fast(){
	document.querySelectorAll("input,select").forEach(a => a.disabled = 1);
	$Id("TITLE1").style.visibility = "visible";
	$Id("SEARCH_TYPE").value = "受入日";
	$Id("new").checked = 1;
	$Id("zan").checked = 1;
	$Id("complete").checked = 1;
	$Id("or").checked = 1;
	$Id("bubun").checked = 1;
	const now = new Date();
	now.setDate(now.getDate() - 7);
	$Id("SEARCH_TXT").type = "date";
	$Id("SEARCH_TXT").value = now.toISOString().split('T')[0];

	$Id("exam1").innerText = "（入力した日付以前/以降での検索が可能）";
	$Id("exam2").innerText = "";
	$Id("exam3").innerText = "";
	document.querySelector('label[for="kan"]').innerHTML = "以前";
	document.querySelector('input[id="bubun"]').checked = 1;
	document.querySelector('label[for="bubun"]').innerHTML = "以降";
	document.querySelector('input[id="non_bubun"]').style.display = "none";
	document.querySelector('label[for="non_bubun"]').style.display = "none";
	prev_search_type = "受入日";
			
	SEARCH_ON("NEW");
	$Id("history_box").style.visibility = "hidden";
	requestAnimationFrame(() => setTimeout(() =>{
		$Id("TITLE1").style.visibility = "hidden";
		document.querySelectorAll("input,select").forEach(a => a.disabled = 0);
		$Id("SEARCH_TXT").focus();
		$Id("TITLE2").style.visibility = "hidden";
	}, 0));
}

function renderNextChunk(Data_out,currentRow,resolve) {
	// INITIAL_ROWSをData_out.lengthすると一括読み込み。
	const INITIAL_ROWS = 200;
	const CHUNK_SIZE = 200;
	const tbody = $Id("R_LIST").tBodies[0];
	const limit = Math.min(currentRow + (currentRow === 0 ? INITIAL_ROWS : CHUNK_SIZE), Data_out.length);
	const progressPercent = $Id('progress-percent');
	
	let htmlChunk = "";
	for (let r = currentRow; r < limit; r++) {
		let row_num = 0;
		Data_out[r][1].forEach(a => a[1].forEach(b => row_num += b[1].length));
		Disp_count += row_num;
		htmlChunk +=`<tr style="border-top-style:solid;border-top-width:2px"><td rowspan="${row_num||""}"><div class="moves">${Data_out[r][0]||""}<br>(${row_num}件)</div></td>`
		Data_out[r][1].forEach((a,j) =>{
			let row_num1 = 0;
			a[1].forEach(b => row_num1 += b[1].length);
			j != 0 ? htmlChunk +=`<tr>` : "";
			htmlChunk +=`<td rowspan=${row_num1}><div class="moves">${a[0]}<br>${row_num != row_num1 ? "("+row_num1+"件)": ""}</div></td>`;
			a[1].forEach((b,i) => {
				const row_num2 = b[1].length;
				i != 0 ? htmlChunk +=`<tr>` : "";
				htmlChunk +=`<td rowspan=${row_num2}><div class="moves" onclick="test_func()">${b[0]}<br>${row_num1 != row_num2 ? "("+row_num2+"件)": ""}</div></td>`;
				b[1].forEach((c,k) => {
					k != 0 ? htmlChunk +=`<tr>` : "";
					let c_tmp = c.slice(0,-2).map(String);
					c_tmp[0] += "<br><span style='font-size:12px'>("+c[c.length-1]+")</font>";
					c_tmp[1] = new Date(Math.round((c[1] - 25569) * 86400 * 1000)).toLocaleDateString('ja-JP', {year: 'numeric',month: '2-digit',day: '2-digit'});
					c_tmp[2] = c_tmp[2].replace(/\s{1,}|[,\/]/g, '<br>');
					c_tmp[3] = c_tmp[3].replace(/ {3,}/g, '<br>');
					c_tmp[4] = c_tmp[4].replace(/\s{3,}/g, '<br>');
					htmlChunk +=`<td><div class="moves">${c_tmp.join("</div></td><td><div class='moves'>").replace("\n","<br>")}</div></td></tr>`;
				});
			});
		});
	}
	const processedCount = currentRow;
	const percent = Math.round((processedCount / (Data_out.length )) * 100);
	progressPercent.textContent = percent;
	currentRow === 0 ? tbody.innerHTML = htmlChunk : tbody.insertAdjacentHTML('beforeend', htmlChunk);
	currentRow = limit;
	currentRow < Data_out.length ? requestAnimationFrame(() => setTimeout(() =>{renderNextChunk(Data_out,currentRow,resolve)}, 0)) : resolve();
}

function enter_key(event){
	// ---テキストボックスフォーカス時にエンターキー押下でフィルタ実行させるための処理 ---
	event.keyCode === 13 ? event.preventDefault() : "";
	event.keyCode === 13 ? SEARCH_ON("NEW") : "";
}

function FuncloadEventListener(){
//  --- イベントリスナー ---
	window.addEventListener("click",() => event.target.id == "L_howto" || (!$Id("tops")?.hidden && event.target.closest("div")?.id != "howto") ? $Id("tops").hidden = !$Id("tops").hidden : "");
	$Id("toTop").addEventListener("click",(e) => window.scrollTo({top: 0}));
	
	$Id("SEARCH_TYPE").addEventListener("change",(e)=>{
		if(["受入日"].includes(prev_search_type)){
			$Id("SEARCH_TXT").value = "";
			document.querySelector('label[for="kan"]').innerHTML = "完全一致";
			document.querySelector('label[for="bubun"]').innerHTML = "含む";
			document.querySelector('input[id="non_bubun"]').style.display = "inline-block";
			document.querySelector('label[for="non_bubun"]').style.display = "inline-block";
		}
		$Id("SEARCH_TXT").type = "text";
		$Id("search").querySelectorAll("input").forEach(a => a.disabled = 0);
		if($Id("SEARCH_TYPE").value == "受入日"){
			$Id("or").disabled =1;
			$Id("and").disabled =1;
			$Id("SEARCH_TXT").type = "date";
			const now = new Date();
			now.setDate(now.getDate() - 7);
			$Id("SEARCH_TXT").value = now.toISOString().split('T')[0];
			$Id("exam1").innerText = "（入力した日付以前/以降での検索が可能）";
			$Id("exam2").innerText = "";
			$Id("exam3").innerText = "";
			document.querySelector('label[for="kan"]').innerHTML = "以前";
			document.querySelector('input[id="bubun"]').checked = 1;
			document.querySelector('label[for="bubun"]').innerHTML = "以降";
			document.querySelector('input[id="non_bubun"]').style.display = "none";
			document.querySelector('label[for="non_bubun"]').style.display = "none";
		}
		else{
			$Id("search").querySelectorAll("input[type='radio']").forEach(a => a.disabled =0);
			$Id("exam1").innerText = "（スペース区切りで複数ワードでの検索可能）";
			$Id("exam2").innerText = "（ワードと完全一致/ワードを含む項目を検索）";
			$Id("exam3").innerText = "（複数ワードでの検索時、いずれか/全てに該当する項目を検索）";
		}
		prev_search_type = e.target.value;
	});
	
	$Id("new").addEventListener("change",(e)=>{
		$Id("history_box").style.visibility = "hidden";
	});
	$Id("cont").addEventListener("change",(e)=>{
		$Id("history_box").style.visibility = "visible";
	});
	$Id("zan").addEventListener("change",(e)=>{
		if($Id("zan").checked || $Id("complete").checked){
			$Id("SEARCH_GO").disabled =0;
			$Id("zan_comp").innerHTML = "（残件分、完了分の表示/非表示の選択）";
		}
		else{
			$Id("SEARCH_GO").disabled =1;
			$Id("zan_comp").innerHTML = "<font color='red'>（少なくともどちらか一方にチェック）</font>";
		}
	});
	$Id("complete").addEventListener("change",(e)=>{
		if($Id("zan").checked || $Id("complete").checked){
			$Id("SEARCH_GO").disabled =0;
			$Id("zan_comp").innerHTML = "（残件分、完了分の表示/非表示の選択）";
		}
		else{
			$Id("SEARCH_GO").disabled =1;
			$Id("zan_comp").innerHTML = "<font color='red'>（少なくともどちらか一方にチェック）</font>";
		}
	});
	const element = document.getElementById('body');
	const indicator = document.getElementById('toTop');
	const checkScroll = () => {
		if (!element || !indicator) return;
		const hasScrollbar = element.scrollHeight > element.clientHeight;
		indicator.style.display = hasScrollbar ? 'block' : 'none';
	};
	const observer = new ResizeObserver(checkScroll);
	observer.observe(element);
	
	window.addEventListener("mousedown",() => {
		if(!event.target.closest("#F_info")){
			$Id("F_info").style.height = "0px";
			$Id("F_info").style.border = "solid 0px #ded1d1";
		}
	});
}

function test_func(){
	const target_div = event.target;
	const target_tr_index = target_div.closest("tr").rowIndex;
	const target_KIBAN = target_div.innerText.split("\n")[0];
	let target_SAKUBAN ;
	if([1,2].includes(target_div.closest("td").cellIndex)){
		target_SAKUBAN = target_div.closest("tr").cells[target_div.closest("td").cellIndex-1].innerText.split(/[ 　]|\n/)[0];
	}
	else{
		let i = target_div.closest("tr").rowIndex-1	;
		while(i>0){
			if($Id("R_LIST").tBodies[0].rows[i].cells.length < 11){
				i--;
			}
			else{
				break;
			}
		}
		target_SAKUBAN = i != 0 ? $Id("R_LIST").tBodies[0].rows[i].cells[0].innerText.split(/[ 　]|\n/)[0] : "";
	}
	console.log(test_out_group2[target_SAKUBAN]);
	console.log(test_out_group2[target_SAKUBAN][target_KIBAN.trim()]);
	let result_MFG = test_out_group2[target_SAKUBAN][target_KIBAN.trim()].map(a => {
		const MFG_Days1 = new Date((a[0] - 25569) * 86400000).toISOString().split('T')[0].replaceAll("-","/");
		const MFG_Days2 = new Date((a[13] - 25569) * 86400000).toISOString().split('T')[0].replaceAll("-","/");
		return [MFG_Days1,...a.slice(1,13),MFG_Days2,...a.slice(14)];
	});

	
	$Id("F_link").innerText = `作番: ${target_SAKUBAN}、基板名: ${target_KIBAN.trim()} の実装情報`
	let HTML_buf = `<thead><tr><th>${test_header.join("</th><th>")}</th></tr>`;
	HTML_buf += `<tbody>`;
	result_MFG.forEach(a => HTML_buf += `<tr><td>${a.join("</td><td>")}</td></tr>`);
	HTML_buf += `</tbody>`;
	document.getElementById("F_TABLE").innerHTML = HTML_buf;
	document.getElementById("F_info").setAttribute("style","height:30%;border-bottom-width:5px;border-top-width:5px;border-right-width:5px;overflow-y:auto");

}
