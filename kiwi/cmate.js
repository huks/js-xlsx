function convert_cmate(workbook) {
	/* Get worksheet */
	var first_sheet_name = workbook.SheetNames[0];
	var ws = workbook.Sheets[first_sheet_name];
	/*
	 * Copy worksheet:
	 * The = operator does not make a copy of the data.
	 * It creates a new reference to the same data.
	 */
	var ws_origin = JSON.parse(JSON.stringify(ws));

	var ws_row_length = ws['!ref'].substr(4);

	work_header(ws);
	// work_cell(ws, "T1", "");
	// work_cell(ws, "U1", "");
	// work_cell(ws, "V1", "");
	ws['!ref'] = "A1:S"+ws_row_length; // UPDATE THE REF

	/* DATA converting */
	var jsonCmate;

	try {
		var pbDataLoaded = JSON.parse(gPbData);

		for(i=2;i<=ws_row_length;i++){
			/* external order number : order number(A) */
			work_cell(ws, "A"+[i], ws_origin["A"+[i]].w);

			/* barcode : WIP */
			work_cell(ws, "B"+[i], NO_DATA);

			/* courier company : logistics companies(O) */
			work_cell(ws, "C"+[i], ws_origin["O"+[i]].w);

			/* recipient name : recipient(I) */
			work_cell(ws, "D"+[i], ws_origin["I"+[i]].w);

			/* recipient id number : WIP */
			work_cell(ws, "E"+[i], NO_DATA);

			/* recipient province : provincial cities and counties(J) */
			var string = ws_origin["J"+[i]].w;
			var strArray = string.split(" ");
			var strProvince = strArray[0];
			var isProvince = strProvince.substring(strProvince.length-1, strProvince.length);
			if (isProvince != "省") {
				// console.log(i + ": " + strProvince + " may not be province");
			}
			work_cell(ws, "F"+[i], strArray[0]);

			/* recipient city : strArray[1] */
			work_cell(ws, "G"+[i], strArray[1]);

			/* recipient county : strArray[2] */
			var strCounty = strArray[2];
			try {
				var isCounty = strCounty.substring(strCounty.length-1, strCounty.length);
				if (isCounty != "区" && isCountry == "省") {
					console.log(i + ": " + strCounty + " may not be country");
				}
			} catch (e) {
				// error
			}		
			work_cell(ws, "H"+[i], strArray[2]);

			/* recipient street and house number : address(K) */
			work_cell(ws, "I"+[i], ws_origin["K"+[i]].w);

			/* recipient contact phone : phone(M) */
			work_cell(ws, "J"+[i], ws_origin["M"+[i]].w);

			/* mail address : zip code(N) */
			work_cell(ws, "K"+[i], ws_origin["N"+[i]].w);

			/* package weight (KG) : WIP */
			work_cell(ws, "L"+[i], NO_DATA);

			/* item unit price (RMB) : price(S) */
			work_cell(ws, "M"+[i], ws_origin["S"+[i]].w);

			/* number of the stuffs : real quantity(U) */
			work_cell(ws, "N"+[i], ws_origin["U"+[i]].w);

			/* payment method : WIP */
			work_cell(ws, "O"+[i], NO_DATA);

			/* product name : platform product name(R) */
			work_cell(ws, "P"+[i], ws_origin["R"+[i]].w);

			/* order generation time : transaction hour(F) */
			work_cell(ws, "Q"+[i], ws_origin["F"+[i]].w);

			/* purchaser platform id : WIP */
			work_cell(ws, "R"+[i], NO_DATA);

			/* rack number : shipment number(P) */
			work_cell(ws, "S"+[i], ws_origin["P"+[i]].w);

			/* empty cells */
			//work_cell(ws, "T"+[i], "");
			//work_cell(ws, "U"+[i], "");
			//work_cell(ws, "V"+[i], "");
		}
	
		/* Display DATA converted in HTML */
		htmlOut(workbook);
	
	} catch (e) {
		alert(e);
	}

	return workbook;	
}