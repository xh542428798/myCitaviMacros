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
// 批量清理所有分组名称，移除空格和冒号等特殊字符，使其更规范。
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
		List<Reference> references = mainForm.GetFilteredReferences();	

		//reference to active Project
		Project activeProject = Program.ActiveProjectShell.Project;
		// Dictionary<string, Category> categoryDictionary = new Dictionary<string, Category>();
		for (int i = 0; i < activeProject.Groups.Count; i++)
		{
			Group mygroup = activeProject.Groups[i];
			//MessageBox.Show(keyword.FullName);
			string group_name = mygroup.FullName;
			// 删除空格
			string trimmedString = group_name.Replace(" ", "");

			// 替换冒号为下划线
			string replacedString = trimmedString.Replace(":", "_");
			replacedString = replacedString.Replace("：", "_");
			
			mygroup.SetValue(GroupPropertyId.Name,replacedString);
			//MessageBox.Show(group_name);
		}
	}
}