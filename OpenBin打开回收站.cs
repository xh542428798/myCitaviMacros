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

// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		Program.ActiveProjectShell.ShowProjectProperties(mainForm);
		Program.ActiveProjectShell.ShowRecycleBinDialog();
		
		
		//Get the active ("primary") MainForm

	}
}