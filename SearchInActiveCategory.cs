using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net;
using System.Linq;
using System.IO;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using SwissAcademic.Citavi.DataExchange;

// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project activeProject = Program.ActiveProjectShell.Project;
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		Category my_Category = mainForm.GetSelectedReferenceEditorCategory();
		//MessageBox.Show(my_group.Name);
		// 在这里修改自己的搜索策略  字段缩写:搜索词
		string search_string = String.Format("rc:\"{0}\" AND t1:请输入", my_Category.FullName);
		//MessageBox.Show(search_string);
		
		Program.ActiveProjectShell.ShowSearchForm(mainForm,SearchFormWorkspace.Extended);
		SearchForm mySearchForm = Program.ActiveProjectShell.SearchForm;
		mySearchForm.SetQuery(search_string);

		
	}
	

}