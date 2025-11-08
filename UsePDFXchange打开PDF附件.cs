using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using System.Diagnostics;
// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		//if this macro should ALWAYS affect all titles in active project, choose:
		//ProjectReferenceCollection references = project.References;		

		//if this macro should affect just filtered rows in the active MainForm, choose:
		List<Reference> references = mainForm.GetSelectedReferences();	

		foreach (Reference reference in references)
		{
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
					string pdfPath = "\"" + path + "\"";
					string newpath = pdfPath.Replace("\\", "/");
					//DebugMacro.WriteLine(newpath);
			        ProcessStartInfo startInfo = new ProcessStartInfo();
			        startInfo.FileName = @"D:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe"; // Adobe Acrobat Reader 的可执行文件路径
			        startInfo.Arguments = newpath;
			        Process.Start(startInfo);
				}
				else
				{
					MessageBox.Show("Please Select PDF location!");
				}
			}


		}
	}
}