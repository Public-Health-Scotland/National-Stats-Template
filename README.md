# How to Produce a Publication Report in Word using RMarkdown

Basic R markdown template for producing summary and full report. 

To download the necessary files, click on Clone or download -> Download ZIP and save it to the folder of your choice. 
![Download zip example](https://github.com/NHS-NSS-transforming-publications/Images/blob/master/RMarkdown7.PNG)
* Go to the zip file, right click on it and choose WinZip -> Extract to here
* You should see the files that are below.
![Folder example](https://github.com/NHS-NSS-transforming-publications/Images/blob/master/RMarkdown_Basic2.png)
	* The ISD-NATIONAL-STATS-REPORT.Rmd file is the RMarkdown file that you open in RStudio, modify as desired, and 	run/knit to create a Word file.  The final Word file will have the exact same name as this file except for the 	extension, which will be .docx.  So in this case, the new file being created will be named ISD-NATIONAL-STATS-REPORT.docx
  * The kitemark_tcm97-17949.jpg file is simply this image which RMarkdown will insert into the final MS Word Publications Report template:
  
  ![Kitemark](https://github.com/NHS-NSS-transforming-publications/RMarkdown_Basic/blob/master/kitemark_tcm97-17949.jpg)
  *	The ISD-NATIONAL-STATS-REPORT_TEMPLATE.docx file is the file in which you set styles for headings and tables.  The RMarkdown file will import the styles, but not the content of this file. You ideally shouldn’t need to modify this since we’ve already set the styles, but this is the file you would change if you wanted to change them.
  *	The Cover_Page.docx file is used to import a custom cover page and a custom footer into the user’s settings.  This file is no longer needed and can be erased after those steps are completed.

## One-time preparation steps for each user
Open the Cover_Page.docx file in Word.
* Save the Cover Page.
  * Press Ctrl + A to select all contents. Go to Insert – Cover Page – Save Selection to Cover Page Gallery. Give it a name (e.g. ISD_Publication_Report) and click OK.
  
	![Cover page example](https://github.com/NHS-NSS-transforming-publications/Images/blob/master/RMarkdown6.PNG)
	
  * Save the footer:
  	* Double click on the footer, and select the whole footer like this by pressing Ctrl + A.
	
	![Footer example](https://github.com/NHS-NSS-transforming-publications/Images/blob/master/RMarkdown2.PNG)
	
	* Then go to Insert – Footer – Save Selection to Footer Gallery. Give it a name (e.g. ISD_Publication_Footer) and click OK.
  * Save the VBA macro:
  	* Go to View – Macros – View Macros. Type the macro name you want to save this as (e.g. SetStyleOfTables) and click Create. It will open up the VBA developer window.
	* Copy the following code to the developer window and click the Save button to save the macro.
	
	```vba
	Sub SetStyleOfAllTablesAndPreserveAlignment()
	 ' SetStyleOfAllTablesAndPreserveAlignment Macro

	     For Each objTable In ActiveDocument.Tables

		 '******This first section is for recording the old column alignments*****'
		 numCols = objTable.Columns.Count 'first find the number of columns within the table.
		 ReDim oldColumnAlignments(numCols) As Integer 'initialize an integer array of length 'numCols'.
		 column_index = 0
		 For Each tableColumn In objTable.Columns
		     oldColumnAlignments(column_index) = tableColumn.Cells(1).Range.ParagraphFormat.Alignment
		     column_index = column_index + 1
		 Next tableColumn
		 '************************************************************************'


		 '-------This section changes the styles of the tables to what they should be. -------'
		 objTable.Style = "ISD_Pubs_Tables"
		 PreviousBookmarkID = objTable.Range.PreviousBookmarkID
		 PreviousBookmarkName = ActiveDocument.Range.Bookmarks(PreviousBookmarkID)

		 If PreviousBookmarkName = "glossary" Then
		     objTable.Style = "Glossary_Style"
		 End If

		 If PreviousBookmarkName = "tableA" Then 'Change these as needed for each style type!
		     objTable.Style = "TableA_Style"
		 End If

		 If PreviousBookmarkName = "tableB" Then
		     objTable.Style = "TableB_Style"
		 End If

		 If PreviousBookmarkName = "tableC" Then
		     objTable.Style = "TableC_Style"
		 End If
		 '------------------------------------------------------------------------------------'


		 '^^^^This last section sets the alignments of each column of the table to what they were ^^^^'
		 '^^^^before the style of the table was changed.^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
		 i = 0
		 For Each tableColumn In objTable.Columns
		     tableColumn.Select
		     Selection.ParagraphFormat.Alignment = oldColumnAlignments(i)
		     i = i + 1
		 Next tableColumn
		 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'

		 objTable.PreferredWidth = 100 'Sets the Preferred Table width to 100% of the width of the page.

	     Next objTable

	 End Sub
	```
