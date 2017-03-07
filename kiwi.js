function work_header(ws) {
	console.log("function.work_header is called");
	work_cell(ws, "A1", "外部订单编号"); // external order number	
	work_cell(ws, "B1", "商品条码"); // barcode
	work_cell(ws, "C1", "快递公司"); // courier company
	work_cell(ws, "D1", "收件人名字"); // recipient name
	work_cell(ws, "E1", "收件人身份证号"); // recipient id number
	work_cell(ws, "F1", "收件人省"); // recipient province
	work_cell(ws, "G1", "收件人市"); // recipient city
	work_cell(ws, "H1", "收件人县区"); // recipient country
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

/* NOT USED */
function work_cell_foo(worksheet, address, value) {
	/*
	 * v: raw value
	 * w: formatted text (if applicable)
	 * 
	 * Built-in export utilities (such as the CSV exporter) will use the w text if it is available.
	 * To change a value, be sure to delete cell.w (or set it to undefined) before attempting to export.
	 * The utilities will regenerate the w text from the number format (cell.z) and the raw value if possible.
	 */
	var w_address = address;
	var w_cell = worksheet[w_address];
	w_cell.v = value;
}

function work_cell(worksheet, address, value) {

	var w_address = address;
	if (worksheet[w_address] == null) {
		worksheet[w_address] = {v:"",t:"s"};
		var w_cell = worksheet[w_address];
		w_cell.v = value;
	} else {
		var w_cell = worksheet[w_address];
		w_cell.t = "s";
		w_cell.v = value;
	}

}

function convert_cmate(workbook) {
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = ws;

	var cmate_row_length = ws['!ref'].substr(4);
	console.log("cmate_row_length: " + cmate_row_length);

	work_header(ws);
	work_cell(ws, "T1", "");
	work_cell(ws, "U1", "");
	work_cell(ws, "V1", "");

	for(i=2;i<=cmate_row_length;i++){
		/* external order number : order number(A) */
		work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

		/* barcode : WIP */
		work_cell(ws, "B"+[i], "BARCODE");

		/* courier company : logistics companies(O) */
		work_cell(ws, "C"+[i], ws_origin["O"+[i]].w);

		/* recipient name : recipient(I) */
		work_cell(ws, "D"+[i], ws_origin["I"+[i]].w);

		/* recipient id number : WIP */
		work_cell(ws, "E"+[i], "RECIPIENT_ID_NUMBER"); 		

		/* recipient province : provincial cities and counties(J) */
		var string = ws_origin["J"+[i]].w;
		var strArray = string.split(" ");
		work_cell(ws, "F"+[i], strArray[0]);

		/* recipient city : strArray[1] */
		work_cell(ws, "G"+[i], strArray[1]);

		/* recipient country : strArray[2] */
		work_cell(ws, "H"+[i], strArray[2]);

		/* recipient recipient street and house number : address(K) */
		work_cell(ws, "I"+[i], ws_origin["K"+[i]].w);

		/* recipient contact phone : phone(M) */
		work_cell(ws, "J"+[i], ws_origin["M"+[i]].w);

		/* mail address : zip code(N) */
		work_cell(ws, "K"+[i], ws_origin["N"+[i]].w);

		/* package weight (KG) : WIP */
		work_cell(ws, "L"+[i], "");

		/* item unit price (RMB) : price(S) */
		work_cell(ws, "M"+[i], ws_origin["S"+[i]].w);

		/* number of the stuffs : real quantity(U) */
		work_cell(ws, "N"+[i], ws_origin["U"+[i]].w);

		/* payment method : WIP */
		work_cell(ws, "O"+[i], "");

		/* product name : platform product name(R) */
		work_cell(ws, "P"+[i], ws_origin["R"+[i]].w);

		/* order generation time : transaction hour(F) */
		work_cell(ws, "Q"+[i], ws_origin["F"+[i]].w);

		/* purchaser platform id : WIP */
		work_cell(ws, "R"+[i], "");

		/* rack number : shipment number(P) */
		work_cell(ws, "S"+[i], ws_origin["P"+[i]].w);

		/* empty cells */
		work_cell(ws, "T"+[i], "");
		work_cell(ws, "U"+[i], "");
		work_cell(ws, "V"+[i], "");
	}	
	return workbook;	
}

function convert_the_get(workbook) {
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = ws;

	var the_get_row_length = ws['!ref'].substr(4);
	console.log("the_get_row_length: " + the_get_row_length);

	work_header(ws);

	for(i=2;i<=the_get_row_length;i++){
		/* external order number : 订单编号 몰 주문번호(A) */
		work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

		/* barcode : WIP */
		work_cell(ws, "B"+[i], "BARCODE");

		/* courier company : WIP */
		work_cell(ws, "C"+[i], "COURIER_COMPANY");

		/* recipient name : 收货人姓名 수화인성명(G) */
		work_cell(ws, "D"+[i], ws_origin["G"+[i]].w);

		/* recipient id number : 买家会员名 회원아이디(D) */
		work_cell(ws, "E"+[i], ws_origin["D"+[i]].w); 		

		/* recipient province : WIP */
		work_cell(ws, "F"+[i], "");

		/* recipient city : WIP */
		work_cell(ws, "G"+[i], "");

		/* recipient country : WIP */
		work_cell(ws, "H"+[i], "");

		/* recipient recipient street and house number : 收货地址 주소(H) */
		work_cell(ws, "I"+[i], ws_origin["H"+[i]].w);

		/* recipient contact phone : 联系手机 연락방식(I) */
		work_cell(ws, "J"+[i], ws_origin["I"+[i]].w);

		/* mail address : WIP */
		work_cell(ws, "K"+[i], "");

		/* package weight (KG) : WIP */
		work_cell(ws, "L"+[i], "");

		/* item unit price (RMB) : WIP */
		work_cell(ws, "M"+[i], "");

		/* number of the stuffs : 购买数量 수량(K) */
		work_cell(ws, "N"+[i], ws_origin["K"+[i]].w);

		/* payment method : WIP */
		work_cell(ws, "O"+[i], "");

		/* product name : 标题 품명(J) */
		work_cell(ws, "P"+[i], ws_origin["J"+[i]].w);

		/* order generation time : WIP */
		work_cell(ws, "Q"+[i], "");

		/* purchaser platform id : WIP */
		work_cell(ws, "R"+[i], "");

		/* rack number : 윈다 송장번호(C) */
		work_cell(ws, "S"+[i], ws_origin["C"+[i]].w);

		/* empty cells */
	}	
	return workbook;	
}