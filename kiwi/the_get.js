function convert_the_get(workbook) {
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = JSON.parse(JSON.stringify(ws));

	var the_get_row_length = ws['!ref'].substr(4);
	// console.log("the_get_row_length: " + the_get_row_length);

	work_header(ws);
	ws['!ref'] = "A1:S"+the_get_row_length;

	var jsonTheGet;

	promiseJSON("db/the_get.json").then(function(response) {
		try {
			jsonTheGet = JSON.parse(response);
		} catch (e) {
			// error
		}

		for(i=2;i<=the_get_row_length;i++){
			/* external order number : 订单编号 몰 주문번호(A) */
			work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

			/* barcode : WIP */
			work_cell(ws, "B"+[i], NO_DATA);

			/* courier company : WIP */
			work_cell(ws, "C"+[i], NO_DATA);

			/* recipient name : 收货人姓名 수화인성명(G) */
			work_cell(ws, "D"+[i], ws_origin["G"+[i]].w);

			/* recipient id number : 买家会员名 회원아이디(D) */
			work_cell(ws, "E"+[i], ws_origin["D"+[i]].w); 		

			/* recipient province : strArray[0] */
			var string = ws_origin["H"+[i]].w;
			var strArray = string.split(" ");
			// console.log(i + ".strArray: " + strArray);
			work_cell(ws, "F"+[i], strArray[0]);

			/* recipient city : strArray[1] */
			work_cell(ws, "G"+[i], strArray[1]);

			/* recipient county : strArray[2] */
			work_cell(ws, "H"+[i], strArray[2]);

			/* recipient street and house number : restAddr */
			// console.log("weee: " + strArray.length);
			var restAddr = strArray[3];
			var length = strArray.length;
			if (length > 4) {
				for (var j=3; j<length; j++) {
					restAddr = restAddr.concat(" "+strArray[j]);
					// console.log("restAddr: " + restAddr);
				}	
			}
			work_cell(ws, "I"+[i], restAddr.substring(0, restAddr.length-8));

			/* recipient contact phone : 联系手机 연락방식(I) */
			work_cell(ws, "J"+[i], ws_origin["I"+[i]].w);

			/* mail address : mailAddr */
			var mailAddr = restAddr.substring(restAddr.length-7, restAddr.length-1);
			work_cell(ws, "K"+[i], mailAddr);

			/* package weight (KG) : WIP */
			work_cell(ws, "L"+[i], NO_DATA);

			/* item unit price (RMB) : WIP */
			work_cell(ws, "M"+[i], NO_DATA);

			/* number of the stuffs : 购买数量 수량(K) */
			work_cell(ws, "N"+[i], ws_origin["K"+[i]].w);

			/* payment method : WIP */
			work_cell(ws, "O"+[i], NO_DATA);

			/* product name : 标题 품명(J) */
			work_cell(ws, "P"+[i], ws_origin["J"+[i]].w);

			/* order generation time : WIP */
			work_cell(ws, "Q"+[i], NO_DATA);

			/* purchaser platform id : WIP */
			work_cell(ws, "R"+[i], NO_DATA);

			/* rack number : 윈다 송장번호(C) */
			work_cell(ws, "S"+[i], ws_origin["C"+[i]].w);

			/* empty cells */
		}

		/* Display DATA converted in HTML */
		htmlOut(workbook);

	}, function(error) {
		console.log("Error!", error);
	});
	
	return workbook;	
}