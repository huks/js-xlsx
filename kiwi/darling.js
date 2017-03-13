function convert_darling(workbook) {
	console.log("function.convert_darling is called");
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = JSON.parse(JSON.stringify(ws));

	var darling_row_length = ws['!ref'].substr(4);

	work_header(ws);
	ws['!ref'] = "A1:S"+darling_row_length;

	/* JSON */
	var jsonDarling;
	loadJSON("db/darling.json", function(response) {
		jsonDarling = JSON.parse(response);
	});

	for(i=2;i<=darling_row_length;i++){
		/* external order number : merchant order(A) */
		work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

		/* barcode : commodity code(J) */
		work_cell(ws, "B"+[i], ws_origin["J"+[i]].w);

		/* courier company : WIP */
		work_cell(ws, "C"+[i], NO_DATA);

		/* recipient name : recipient(C) */
		work_cell(ws, "D"+[i], ws_origin["C"+[i]].w);

		/* recipient id number : id number(D) */
		work_cell(ws, "E"+[i], ws_origin["D"+[i]].w); 		

		/* recipient province : strArray[0] */
		var string = ws_origin["E"+[i]].w;
		var strArray = string.split(" ");
		work_cell(ws, "F"+[i], strArray[0]);

		/* recipient city : strArray[1] */
		work_cell(ws, "G"+[i], strArray[1]);

		/* recipient county : strArray[2] */
		work_cell(ws, "H"+[i], strArray[2]);

		/* recipient street and house number : strArray[3] */
		work_cell(ws, "I"+[i], strArray[3]);

		/* recipient contact phone : recipient phone(F) */
		work_cell(ws, "J"+[i], ws_origin["F"+[i]].w);

		/* mail address : WIP */
		work_cell(ws, "K"+[i], NO_DATA);

		/* package weight (KG) : WIP */
		work_cell(ws, "L"+[i], NO_DATA);

		/* item unit price (RMB) : unit price(M) */		
		try {
			var fooPrice = getPriceByBarcode(jsonDarling, ws_origin["J"+[i]].w)
			work_cell(ws, "M"+[i], fooPrice[0].price);
		} catch (e) {
			work_cell(ws, "M"+[i], "UNDEFINED");
		}

		/* number of the stuffs : quantity(L) */
		work_cell(ws, "N"+[i], ws_origin["L"+[i]].w);

		/* payment method : WIP */
		work_cell(ws, "O"+[i], NO_DATA);

		/* product name : name(K) */
		work_cell(ws, "P"+[i], ws_origin["K"+[i]].w);

		/* order generation time : WIP */
		work_cell(ws, "Q"+[i], NO_DATA);

		/* purchaser platform id : WIP */
		work_cell(ws, "R"+[i], NO_DATA);

		/* rack number : waybill number(B) */
		work_cell(ws, "S"+[i], ws_origin["B"+[i]].w);

		/* empty cells */
	}	
	return workbook;	
}