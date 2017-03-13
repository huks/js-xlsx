function convert_beauty(workbook) {
	// console.log("function.convert_beauty is called");
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = JSON.parse(JSON.stringify(ws));

	var beauty_row_length = ws['!ref'].substr(4);

	work_header(ws);
	ws['!ref'] = "A1:S"+beauty_row_length;

	/* JSON */
	var jsonBeauty;
	// loadJSON("db/beauty.json", function(response) {
	// 	jsonBeauty = JSON.parse(response);
	// });

	promiseJSON("db/beauty.json").then(function(response) {
		jsonBeauty = JSON.parse(response);

		for(i=2;i<=beauty_row_length;i++){
			/* external order number : merchant order(A) */
			work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

			/* barcode : commodity code(I) */
			work_cell(ws, "B"+[i], ws_origin["I"+[i]].w);

			/* courier company : WIP */
			work_cell(ws, "C"+[i], NO_DATA);

			/* recipient name : recipient(C) */
			work_cell(ws, "D"+[i], ws_origin["C"+[i]].w);

			/* recipient id number : WIP */
			work_cell(ws, "E"+[i], NO_DATA); 		

			/* recipient province : WIP */
			work_cell(ws, "F"+[i], NO_DATA);

			/* recipient city : WIP */
			work_cell(ws, "G"+[i], NO_DATA);

			/* recipient county : WIP */
			work_cell(ws, "H"+[i], NO_DATA);

			/* recipient street and house number : recipient address(D) */
			work_cell(ws, "I"+[i], ws_origin["D"+[i]].w);

			/* recipient contact phone : recipient phone(E) */
			work_cell(ws, "J"+[i], ws_origin["E"+[i]].w);

			/* mail address : WIP */
			work_cell(ws, "K"+[i], NO_DATA);

			/* package weight (KG) : WIP */
			work_cell(ws, "L"+[i], NO_DATA);

			/* item unit price (RMB) : unit price(L) */		
			try {
				var fooPrice = getPriceByBarcode(jsonBeauty, ws_origin["I"+[i]].w)
				work_cell(ws, "M"+[i], fooPrice[0].price);
			} catch (e) {
				work_cell(ws, "M"+[i], "UNDEFINED");
			}

			/* number of the stuffs : quantity(K) */
			work_cell(ws, "N"+[i], ws_origin["K"+[i]].w);

			/* payment method : WIP */
			work_cell(ws, "O"+[i], NO_DATA);

			/* product name : name(J) */
			work_cell(ws, "P"+[i], ws_origin["J"+[i]].w);

			/* order generation time : WIP */
			work_cell(ws, "Q"+[i], NO_DATA);

			/* purchaser platform id : WIP */
			work_cell(ws, "R"+[i], NO_DATA);

			/* rack number : waybill number(B) */
			work_cell(ws, "S"+[i], ws_origin["B"+[i]].w);

			/* empty cells */
		}

		/* Display DATA converted in HTML */
		htmlOut(workbook);

	}, function(error) {
		console.log("Error!", error);
	});		

	return workbook;	
}