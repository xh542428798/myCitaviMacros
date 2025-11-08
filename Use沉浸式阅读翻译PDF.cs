using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using SwissAcademic;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using System.Diagnostics;
public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		List<Location> locations = mainForm.GetSelectedElectronicLocations();
		if (locations.Count !=1)
		{
			MessageBox.Show("Please Select Only One PDF location!");
		}else
		{
			Location location = locations[0];
			string path = location.Address.Resolve().LocalPath;
			DebugMacro.WriteLine(location.Address.Resolve().LocalPath);
			if (path.EndsWith(".pdf"))
			{
				
				string newpath = path.Replace("\\", "/");
				string command = "start msedge \"chrome-extension://amkbmndfnliijdhojkpoglbnaaahippg/pdf/index.html?file=file:///"+newpath +"\"";	
				//string command = "start chrome \"chrome-extension://bpoadfkcbjbfhfodiogcnhhhpibjhbnh/pdf/index.html?file=file:///"+newpath +"\"";
				DebugMacro.WriteLine(command);
				Process.Start("cmd.exe", "/c " + command);
			}
			else
			{
				MessageBox.Show("Please Select PDF location!");
			}
		}
		


	}
}