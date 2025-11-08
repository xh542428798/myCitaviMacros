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

// 将1_SetTitle5WordToCategoryAndKnowledge.cs中创建的多个Knowledge组合并为一个组
// 此宏目前作用不大了
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

		var categories = project.Categories.ToList();

        //MergeItemsDialog mergeDialog = new SwissAcademic.Citavi.Shell.MergeItemsDialog;
		List<Category> category_first = new List<Category> { categories[0] }; // categories.Take(2).ToList(); 
		foreach (Category category in categories.Skip(1))
		{
			// MessageBox.Show(category.Name);
			List<KnowledgeItem> knowledgeItems = category.KnowledgeItems.ToList();
			foreach (KnowledgeItem knowledgeItem in knowledgeItems)
			{
				//MessageBox.Show(knowledgeItem.FullName);
				//knowledgeItem.Categories.Clear();
				knowledgeItem.Categories.AddRange(category_first);
				// project.Categories.Remove(category);
			}
		}
	}
}