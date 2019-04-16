# How to Produce a Publication Report in Word using RMarkdown

Basic R markdown template for producing summary and full report. 

To download the necessary files, click on Clone or download -> Download ZIP and save it to the folder of your choice. 
![Download zip example](https://github.com/NHS-NSS-transforming-publications/Images/blob/master/RMarkdown7.PNG)
* Go to the zip file, right click on it and choose WinZip -> Extract to here
* You should see the files that are below.
![Folder example](https://github.com/NHS-NSS-transforming-publications/Images/blob/master/RMarkdown_Basic2.png)
	* The ISD-NATIONAL-STATS-REPORT.Rmd file is the RMarkdown file that you open in RStudio, modify as desired, and 	run/knit to create a Word file.  The final Word file will have the exact same name as this file except for the 	extension, which will be .docx.  So in this case, the new file being created will be named ISD-NATIONAL-STATS-REPORT.docx
  * The kitemark_tcm97-17949.jpg file is simply this image which RMarkdown will insert into the final MS Word Publications Report template:
  ![](https://github.com/NHS-NSS-transforming-publications/RMarkdown_Basic/blob/master/kitemark_tcm97-17949.jpg)
  *	The ISD-NATIONAL-STATS-REPORT_TEMPLATE.docx file is the file in which you set styles for headings and tables.  The RMarkdown file will import the styles, but not the content of this file. You ideally shouldn’t need to modify this since we’ve already set the styles, but this is the file you would change if you wanted to change them.
  *	The Cover_Page.docx file is used to import a custom cover page and a custom footer into the user’s settings.  This file is no longer needed and can be erased after those steps are completed.

## One-time preparation steps for each user
Open the Cover_Page.docx file in Word.
  *	Save the Cover Page.
    *	Press Ctrl + A to select all contents. Go to Insert – Cover Page – Save Selection to Cover Page Gallery. Give it a name (e.g. ISD_Publication_Report) and click OK.
![Cover page example](https://github.com/NHS-NSS-transforming-publications/Images/blob/master/RMarkdown_Basic2.png)
  
