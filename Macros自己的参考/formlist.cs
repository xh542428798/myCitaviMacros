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
		
		List<string> formNameList = new List<string>();
		for (int i = 0; i < Application.OpenForms.Count; i++)
        {
			Form myform = Application.OpenForms[i];
			if (!myform.Visible) continue;
			formNameList.Add(myform.Name);
			//form.Activate();
			//MessageBox.Show(myform.Name);
		}
		
		// Get the active form
		Form myActivaform = Program.ActiveProjectShell.ActiveForm;
		//MessageBox.Show(myActivaform.Name);
		// Set the next form to be activated
		int currentIndex = formNameList.IndexOf(myActivaform.Name);
		int nextIndex = (currentIndex + 1) % formNameList.Count; // Get the index of the next form in a circular manner
		string nextFormName = formNameList[nextIndex];
		
		// Activate the next form
		Form nextForm = Application.OpenForms[nextFormName];
		//MessageBox.Show(nextForm.Name);
		nextForm.Activate();
	}
}