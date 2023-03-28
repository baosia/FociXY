//Before executing, select cells and add them to the ROI manager from the DAPI chanel image
//Loop through chanel images to look for green chanel
for (i=1; i<=nImages(); i++) {
    selectImage(i);
    current = getTitle();
    if (matches(current, ".*C=1")){
    	run("8-bit");
    	//Loop through ROIs manually selected from DAPI chanel image
    	n = roiManager('count');
			for (j = 0; j < n; j++) {
    			roiManager('select', j);
    			// process roi here
    			//Change prominence value as needed
			run("Find Maxima...", "prominence=10 output=[List]");
    			//Below adds an "ROI" row to the results output
    			IJ.renameResults("Results");
    			for (row=0; row<nResults; row++) {
    				setResult("ROI", row, j);
    			}
    			//Send results to an excel file (requires read and write excel plugin)
    			selectWindow("Results");
    			//specify path to your excel sheet destination
			run("Read and Write Excel", "file=[C:/Users/"USER"/Documents/fociXY.xlsx] sheet=Green");
			} 
    }   
    	else {
    		//Loop through chanels again to find the red chanel, and repeat block above.
    		if (matches(current, ".*C=2")){
    		run("8-bit");
    		m = roiManager('count');
				for (k = 0; k < m; k++) {
    				roiManager('select', k);
    				//Change prominence value as needed
				run("Find Maxima...", "prominence=20 output=[List]");
    				IJ.renameResults("Results");
    				for (row=0; row<nResults; row++) {
    					setResult("ROI", row, k);
    				}
    				selectWindow("Results");
				//Specify same path as above (will write results to a second sheet)
    				run("Read and Write Excel", "file=[C:/Users/"USER"/Documents/fociXY.xlsx] sheet=Red");
				}
    		}
	 }
}
