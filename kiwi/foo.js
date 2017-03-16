var NO_DATA = "n/a";

function htmlOut(data) {
	var output = JSON.stringify(to_json(data), 2, 2);

	if (out.innerText === undefined) {
		// console.log("out.textContent:");
		out.textContent = output;
	} else {
		// console.log("out.innerText:");
		out.innerText = output;
	} 
	if (typeof console !== 'undefined') {
		console.log("output", new Date());
	} 
}

function getPriceByBarcode(json, brcd) {
	return json.filter(function(item){
		return item.barcode.replace(/\s/g, '') == brcd; 
	});
}

/* Better use this function instead of loadJSON() */
function promiseJSON(url) {
	console.log("promiseJSON() is called");
	// Return a new promise.
	return new Promise(function(resolve, reject) {
		// Do the usual XHR stuff
		var req = new XMLHttpRequest();
		req.overrideMimeType("application/json");
		req.open("GET", url);

		req.onload = function() {
			if (req.status == 200) {
				// Resolve the promise with the response text
				// console.log("Resolve the promise...!");
				resolve(req.response);
			}
			else {
				// Otherwise reject with the status text
				// which will hopefully be a meaningful error
				reject(Error(req.statusText));
			}
		};

		// Handle network errors
		req.onerror = function() {
			reject(Error("Network Error"));
		};

		// Make the request
		req.send();
	});
}

function loadJSON(url, callback) {
	var req = new XMLHttpRequest();
	req.overrideMimeType("application/json");
	req.open('GET', url, false); // This is not actually recommended, as it has to wait for the 'server' response.
	req.onreadystatechange = function () {
		if (req.readyState == 4 && req.status == "200") {
			// Required use of an anonymous callback as .open will NOT return a value but simply returns undefined in asynchronous mode
			callback(req.responseText);
		}
	};
	req.send(null);  
}

function work_header(ws) {
	work_cell(ws, "A1", "外部订单编号"); // external order number	
	work_cell(ws, "B1", "商品条码"); // barcode
	work_cell(ws, "C1", "快递公司"); // courier company
	work_cell(ws, "D1", "收件人名字"); // recipient name
	work_cell(ws, "E1", "收件人身份证号"); // recipient id number
	work_cell(ws, "F1", "收件人省"); // recipient province
	work_cell(ws, "G1", "收件人市"); // recipient city
	work_cell(ws, "H1", "收件人县区"); // recipient county
	work_cell(ws, "I1", "收件人街道及门牌号"); // recipient street and house number
	work_cell(ws, "J1", "收件人联系电话"); // recipient contact phone
	work_cell(ws, "K1", "邮件地址"); // mail address	
	work_cell(ws, "L1", "包裹重量（KG）"); // package weight (KG)
	work_cell(ws, "M1", "物品单价（RMB）"); // item unit price (RMB)
	work_cell(ws, "N1", "物品数量"); // number of the stuffs
	work_cell(ws, "O1", "支付方式"); // payment method
	work_cell(ws, "P1", "商品名称"); // product name
	work_cell(ws, "Q1", "订单生成时间（年-月-日 时:分:秒）"); // order generation time
	work_cell(ws, "R1", "购买人平台ID"); // purchaser platform id
	work_cell(ws, "S1", "RACK NUMBER"); // rack number
}

function work_cell(worksheet, address, value) {

	var w_cell;

	if (worksheet[address] == null) {
		worksheet[address] = {t:"s",v:"",w:""};
		w_cell = worksheet[address];
		w_cell.v = value;
		w_cell.w = value;
	} else {
		w_cell = worksheet[address];
		w_cell.t = "s";
		w_cell.v = value;
		w_cell.w = value;
	}

	// console.log("["+address+"]: "+JSON.stringify(w_cell));

}