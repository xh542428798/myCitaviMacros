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
			//string fieldName = "SeriesTitle";
			//string property = reference.GetType().GetProperty(fieldName).GetValue(reference).ToString();
				// Locations
			//foreach(Location location in reference.Locations)
			//{
			//	bool yes = reference.Locations.Contains(references[1].Locations[0]);
			//	MessageBox.Show(yes.ToString());
			//}

			//reference.GetType().GetProperty(fieldName).SetValue(reference,"IF:455");
			//property = reference.GetType().GetProperty(fieldName).GetValue(reference).ToString();
			//MessageBox.Show(property);
			// your code
			List<Person> refPersons = reference.Authors.ToList();
			List<string> nameString = new List<string>(); // 创建一个空的 List<string>
			foreach (Person person in refPersons)
			{
				if (reference.Authors == null) continue;
				nameString.Add(person.FullName);
				
			}
			        // 将 List 转换为数组
        	string[] strings = nameString.ToArray();
			Array.Sort(strings);
			if (strings.Length > 0)
			{
				string result = string.Join("\n", strings); // 如果有两个以上的字符串，用换行符连接它们
				MessageBox.Show(result);
			};
			
		}
	}
}