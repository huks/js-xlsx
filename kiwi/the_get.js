function convert_the_get(workbook) {
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = JSON.parse(JSON.stringify(ws));

	var the_get_row_length = ws['!ref'].substr(4);
	console.log("the_get_row_length: " + the_get_row_length);

	work_header(ws);
	ws['!ref'] = "A1:S"+the_get_row_length;

	for(i=2;i<=the_get_row_length;i++){
		/* external order number : 订单编号 몰 주문번호(A) */
		work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

		/* barcode : WIP */
		work_cell(ws, "B"+[i], "BARCODE");

		/* courier company : WIP */
		work_cell(ws, "C"+[i], "");

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