function convert_childking(workbook) {
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/* Copy worksheet */
	var ws_origin = JSON.parse(JSON.stringify(ws));

	var childking_row_length = ws['!ref'].substr(4);
	// console.log("childking_row_length: " + childking_row_length);

	work_header(ws);
	ws['!ref'] = "A1:S"+childking_row_length;

	/* JSON */
	var jsonChildking;
	loadJSON("db/childking.json", function(response) {
		jsonChildking = JSON.parse(response);
	});

	for(i=2;i<=childking_row_length;i++){
		/* external order number : merchant order(A) */
		work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

		/* barcode : commodity code(N) */
		work_cell(ws, "B"+[i], ws_origin["N"+[i]].w);

		/* courier company : WIP */
		work_cell(ws, "C"+[i], NO_DATA);

		/* recipient name : recipient(C) */
		work_cell(ws, "D"+[i], ws_origin["C"+[i]].w);

		/* recipient id number : id number(D) */
		work_cell(ws, "E"+[i], ws_origin["D"+[i]].w); 		

		/* recipient province : recipient province(E) */
		work_cell(ws, "F"+[i], ws_origin["E"+[i]].w);

		/* recipient city : recipient city(F) */
		work_cell(ws, "G"+[i], ws_origin["F"+[i]].w);

		/* recipient country : recipient area(G) */
		work_cell(ws, "H"+[i], ws_origin["G"+[i]].w);

		/* recipient street and house number : shipping address(H) */
		work_cell(ws, "I"+[i], ws_origin["H"+[i]].w);

		/* recipient contact phone : recipient phone(J) */
		work_cell(ws, "J"+[i], ws_origin["J"+[i]].w);

		/* mail address : WIP */
		work_cell(ws, "K"+[i], NO_DATA);

		/* package weight (KG) : single weight(S) */
		work_cell(ws, "L"+[i], ws_origin["S"+[i]].w);

		/* item unit price (RMB) : unit price(U) */		
		try {
			var fooPrice = getPriceByBarcode(jsonChildking, ws_origin["N"+[i]].w)
			work_cell(ws, "M"+[i], fooPrice[0].price);
		} catch (e) {
			work_cell(ws, "M"+[i], "UNDEFINED");
		}

		/* number of the stuffs : quantity(T) */
		work_cell(ws, "N"+[i], ws_origin["T"+[i]].w);

		/* payment method : WIP */
		work_cell(ws, "O"+[i], NO_DATA);

		/* product name : name(P) */
		work_cell(ws, "P"+[i], ws_origin["P"+[i]].w);

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