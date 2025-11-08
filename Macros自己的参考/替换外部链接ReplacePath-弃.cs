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
// location绝对路径修改
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
			// your code
			// DebugMacro.WriteLine(reference.Doi);
			List<Location> refLocation= reference.Locations.ToList();
            foreach (Location location in refLocation)
            {
	            if (reference.Locations == null) continue;
				if (string.IsNullOrEmpty(location.Address.ToString())) continue;
				if (location.LocationType != LocationType.ElectronicAddress) continue;
				string filePath = location.Address.ToString();
				if (filePath.Contains("OneDrive - 中山大学"))
				{
					// DebugMacro.WriteLine(filePath);
					location.Address.ReplaceTextInPaths(findText: "OneDrive - 中山大学", replacementText:"OneDrive - xiehui1573"); // 修改路径名，替换里面的目标
					DebugMacro.WriteLine(location.Address.ToString());
					// DebugMacro.WriteLine(location.LocationType.ToString());
					//reference.Doi = filePath.Replace("https://doi.org/", "");
				}
				// DebugMacro.WriteLine(reference.Doi);

			}
		}
	}
}