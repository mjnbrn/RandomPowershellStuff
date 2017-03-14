using System;
using System.Windows.Forms;
using System.Management.Automation;
using System.Collections.ObjectModel;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string file = "";
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
            }
            //MessageBox.Show(file);

            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                string excelCSV = @"Function ExcelCSV ($File){
                $excelFile = $File;
                $Excel = New-Object -ComObject Excel.Application;
                $Excel.Visible = $false;
                $Excel.DisplayAlerts = $false;
                $wb = $Excel.Workbooks.Open($excelFile);
                foreach ($ws in $wb.Worksheets){$ws.SaveAs($File + '_' + $ws.name + '.csv', 6);}
                $Excel.Quit();}";

                string PSRaffle = @"
                Function Raffle ($InputFile) {
                ExcelCSV $InputFile;
                $Entries = Import-Csv ($InputFile+""_Raffle.csv"");
                $Tickets = @{ };
                $TicketNo = 0;

                foreach ($Row in $Entries){

                    for ($i = 1; $i -lt $row.Tickets; $i++){
                $Tickets.$TicketNo = $Row.Name;
                $TicketNo++;
                    }
                }
                $First = Get-Random -Maximum $TicketNo;
                $Second = Get-Random -Maximum $TicketNo;
                while ($Second-eq $First)
                {
                    $Second = Get-Random -Maximum $TicketNo;
                }
                $Third = Get-Random -Maximum $TicketNo;
                while ($Third -eq $First -or $Third -eq $Second)
                {
                    $Third = Get-Random -Maximum $TicketNo;
                }
                return ""First place is $($Tickets[$First]) with Ticket number $($First)`r`nSecond place is $($Tickets[$Second]) with Ticket number $($Second)`r`nThird place is $($Tickets[$Third]) with Ticket number $($Third)"";
                }
            ";

                // use "AddScript" to add the contents of a script file to the end of the execution pipeline.
                // use "AddCommand" to add individual commands/cmdlets to the end of the execution pipeline.
                PowerShellInstance.AddScript(excelCSV,false);
                PowerShellInstance.AddScript(PSRaffle, false);
                PowerShellInstance.Invoke();
                PowerShellInstance.Commands.Clear();
                PowerShellInstance.AddCommand("Raffle").AddParameter("InputFile", file);
                Collection<PSObject> PSOutput = PowerShellInstance.Invoke();
                foreach (PSObject outputItem in PSOutput)
                {
                    // if null object was dumped to the pipeline during the script then a null
                    // object may be present here. check for null to prevent potential NRE.
                    if (outputItem != null)
                    {
                        //TODO: do something with the output item 
                        //Console.WriteLine(outputItem.BaseObject.GetType().FullName);
                        MessageBox.Show(outputItem.BaseObject.ToString(), "Raffle Winners!");
                        //MessageBox.Show(outputItem.BaseObject.ToString() + "\n");
                    }
                }


                //MessageBox.Show("We Done!");

            }


        }
    }
}
