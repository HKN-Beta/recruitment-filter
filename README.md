1. Run Cognos 'record copy - PUID entry' report for all ECE students. (Note that these must be processed in batches, as Cognos has a limit of 1000 entries at once per report.)
2. When Cognos reports are done in batches, merge the resulting excel reports. Rename this file to "HKN [SEM YEAR] Cognos" in the HKN folder.

	A. Select the worksheets that you want to move or copy. If you want to move or copy more than one, press and hold the Ctrl key, and then click the tabs of the sheets you want to copy.
	If you want to select all sheets, right-click a sheet tab, and then click Select All Sheets. Note: When multiple worksheets are selected, [Group] appears in the title bar at the top of the worksheet.
	To cancel a selection of multiple worksheets in a workbook, click any unselected worksheet.
	If no unselected sheet is visible, right-click the sheet tab of a selected sheet, and then click Ungroup Sheets on the shortcut menu. 
	
	B. Select Home > Format > Move or Copy Sheet. Or, you can also right-click a selected sheet tab, and then click Move or Copy.

	C. In the Move or Copy dialog box, click the workbook to which you want to move or copy the selected sheets.
	Or you can click new book to move or copy the selected sheets to a new workbook.
	
	D. In the Before sheet list, click the sheet that should be after the moved or copied sheet. Or you can click (move to end) to insert after the last sheet in the workbook.
	
	E. To copy the sheets instead of moving them, in the Move or Copy dialog box, select the Create a copy check box.
	When you create a copy of the worksheet, the worksheet is duplicated in the destination workbook.
	When you move a worksheet, the worksheet is removed from the original workbook and appears in the destination workbook only.
	
	Notes and tips:
	• To rename the moved or copied worksheet in the destination workbook, right-click its sheet tab, click Rename, and then type the new name in the sheet tab.
	• To change the color of the sheet tab, right-click the sheet tab, click Tab Color, and then click the color that you want to use.
	• Worksheets that you move or copy to another workbook will use the theme fonts, colors, and effects that are applied to the destination workbook.

3. Delete the previously generated output file. This will be named "HKN_list.xlst"

4. Run your startup file in MATLAB.

5. Open the HKN folder in MATLAB. Enter: HKN('File Name') 
Note: this will take quite a while (1-2 hours) to run.

6. Verify that the last row in the output file has the lowest GPA.

7. Clean the output file (each tab corresponding to different class standings) so that only the names of the eligible students and their major remain.
Alphabetize the file. Send to the HKN faculty or staff advisor. Sometimes the minimum number of completed ECE credits and the percentile for each classification level,
so make sure you check with the HKN contact about these numbers. Spring 2018 was minimum 10 credits ECE, top 20% of sophomores, top 25% juniors, top 30% seniors.
Percentile can be calculated in Excel by typing =PERCENTILE.EXC(datarange:datarange, percentile in decimal). Send the cut-off GPAs with the names.