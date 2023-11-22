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
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		//if this macro should ALWAYS affect all titles in active project, choose:
		//ProjectReferenceCollection references = project.References;		

		//if this macro should affect just filtered rows in the active MainForm, choose:
		List<Reference> references = mainForm.GetSelectedReferences();	
		
		foreach (Reference reference in references)
		{
	        List<Location> refLocation = reference.Locations.ToList();
            foreach (Location location in refLocation)
            {
                if (reference.Locations == null) continue;
				if (string.IsNullOrEmpty(location.Address.ToString())) continue;
				if (location.LocationType != LocationType.ElectronicAddress) continue;
				string filePath = location.Address.ToString();
				//reference.Locations.Remove
				
				if (filePath.StartsWith(@"\\")) //(filePath.EndsWith(".pdf") && filePath.StartsWith(@"\\"))
				{
					DebugMacro.WriteLine(filePath);
				    filePath = "G:" + filePath;
					DebugMacro.WriteLine(filePath);
					Location newlocation = new Location(reference, LocationType.ElectronicAddress, filePath);
        			reference.Locations.Add(newlocation);
					reference.Locations.Remove(location);
				}
				// 判断路径,如果字符串包含分号，则将其拆分成多个路径并打印出来。否则，将打印原始字符串。
				if (filePath.Contains(";"))
				{
				    string[] paths = filePath.Split(';');
				    foreach (string path in paths)
				    {	
						//path = path.Trim();
						DebugMacro.WriteLine(path);
						Location newlocation = new Location(reference, LocationType.ElectronicAddress, path.Trim());
        				reference.Locations.Add(newlocation);  
				    }
					reference.Locations.Remove(location);
				}
				
				//MessageBox.Show(path);
				
                
            }
		}
	}
}