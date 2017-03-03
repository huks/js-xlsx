function work_cell(worksheet, address, value) {
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

/* NOT USED */
// function format_text_cell(worksheet, address, value) {
// 	var f_address = address;
// 	var f_cell = worksheet[f_address];
// 	f_cell.w = value;
// }

function convert_cmate(workbook) {
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = ws;

	var wb_row_length = ws['!ref'].substr(4);
	console.log("wb_row_length: " + wb_row_length);

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
	work_cell(ws, "T1", "");
	work_cell(ws, "U1", "");
	work_cell(ws, "V1", "");

	for(i=2;i<=wb_row_length;i++){
		/* external order number : order number(A1) OR original order number(E1) */
		work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);
		/* barcode : WIP */
		work_cell(ws, "B"+[i], "BARCODE");
		/* courier company : logistics companies(O1) */
		work_cell(ws, "C"+[i], ws_origin["O"+[i]].w);
		/* recipient name : recipient(I1) */
		work_cell(ws, "D"+[i], ws_origin["I"+[i]].w);
		/* recipient id number : WIP */
		work_cell(ws, "E"+[i], "RECIPIENT_ID_NUMBER"); 		
		/* recipient province : provincial cities and counties(J1) */
		var string = ws_origin["J"+[i]].w;
		var strArray = string.split(" ");
		work_cell(ws, "F"+[i], strArray[0]);
		/* recipient city : WIP */
		work_cell(ws, "G"+[i], strArray[1]);
		/* recipient country : WIP */
		work_cell(ws, "H"+[i], strArray[2]);
		/* recipient recipient street and house number : address(K1) */
		work_cell(ws, "I"+[i], ws_origin["K"+[i]].w);
		/* recipient contact phone : phone(M1) */
		work_cell(ws, "J"+[i], ws_origin["M"+[i]].w);
		/* mail address : zip code(N1) */
		work_cell(ws, "K"+[i], ws_origin["N"+[i]].w);
		/* package weight (KG) : WIP */
		work_cell(ws, "L"+[i], "");
		/* item unit price (RMB) : price(S1) */
		work_cell(ws, "M"+[i], ws_origin["S"+[i]].w);
		/* number of the stuffs : real quantity(U1) */
		work_cell(ws, "N"+[i], ws_origin["U"+[i]].w);
		/* payment method : WIP */
		work_cell(ws, "O"+[i], "");
		/* product name : platform product name(R1) */
		work_cell(ws, "P"+[i], ws_origin["R"+[i]].w);
		/* order generation time : transaction time(F1) */
		work_cell(ws, "Q"+[i], ws_origin["F"+[i]].w);
		/* purchaser platform id : business code(Q1) */
		work_cell(ws, "R"+[i], ws_origin["Q"+[i]].w);
		/* rack number : shipment number(P1) */
		work_cell(ws, "S"+[i], ws_origin["P"+[i]].w);
		/* empty cells */
		work_cell(ws, "T"+[i], "");
		work_cell(ws, "U"+[i], "");
		work_cell(ws, "V"+[i], "");
	}	

	return workbook;	
}