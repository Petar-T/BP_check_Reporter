# BP_Check-To-Excel & BP_Check-Reporter
<ul>
  <li><B>BP_Check_To_Excel</B> os Powershell utility to run BP_Check (and any simmilar) and get results into Excel. <BR>Excel is created in user's temporary folder , and all plans detected during execution are extracted to physical files so they can be opened in SSMS during analysis </li>
  <li><B>BP_Check_Reporter</B> is Powershell script to drill into BP-check results file(s), it also provides PowerBi file  </li>
</ul> 
**Requirements** :
<ul>
  <li>To run <B>BP_Check_To_Excel</B>  ImportExcel Module is required, downloadable from https://www.powershellgallery.com/packages/ImportExcel   </li>
  <li> To excute <B>BP_Check_Reporter</B> you need output of bp_check script from SQL server in format Servername.rpt. There might be unlimited outputs in same batch   </li>
</ul>

**Result</B>**: <BR>results.csv and correlating PowerBI report  <BR>

**Process** 
Run Poweshell script , it will create results.csv, change source on PowerBI file , that is it!

<BR>
![image](https://github.com/Petar-T/BP_check_Reporter/assets/47648550/506eb4a0-c2fb-4503-9705-2d8ad039e755?raw=true "Image 01")
<BR>
![PowerBI demo page](https://github.com/Petar-T/BP_check_Reporter/assets/47648550/4651cd60-29be-48e0-bab0-499157b4732c)
