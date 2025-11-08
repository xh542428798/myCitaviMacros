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
// 针对一个catrgory组，批量打印文本到output中（DebugMacro.WriteLine），将文字复制到md中，就是符合要求的素材文档1，由于使用citavi，实际上素材文档1相当于在citavi里
// 因此可以直接制作2.1
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
		// var knowledgeItems = mainForm.GetSelectedKnowledge();
        var category = mainForm.GetSelectedKnowledgeOrganizerCategory();

        var knowledgeItems = category.KnowledgeItems.ToList();
		
		string relateSubheading="";
		
		foreach (KnowledgeItem knowledgeitem in knowledgeItems)
		{
			Reference reference = knowledgeitem.Reference;
			// MessageBox.Show(knowledgeitem.QuotationType.ToString());
			if ( knowledgeitem.KnowledgeItemType==KnowledgeItemType.Subheading)
			{
				relateSubheading = knowledgeitem.CoreStatement;
				DebugMacro.WriteLine("## "+knowledgeitem.CoreStatement);
			}else
			{
				// 素材文档1
				//DebugMacro.WriteLine("## "+knowledgeitem.CoreStatement);
				//string path = knowledgeitem.Address.DataContractFullUriString;			
		        //path = path.Replace("\\", "/");// Replace backslashes with forward slashes
		        //path = path.Replace(" ", "%20");// Replace spaces with %20
				//DebugMacro.WriteLine("![]("+path+")");
				
				// 素材文档2.1
				DebugMacro.WriteLine(knowledgeitem.CoreStatement);
				DebugMacro.WriteLine("### "+knowledgeitem.CoreStatement);
				//// Subheading
				DebugMacro.WriteLine(relateSubheading);
				string path = knowledgeitem.Address.DataContractFullUriString;			
		        path = path.Replace("\\", "/");// Replace backslashes with forward slashes
		        path = path.Replace(" ", "%20");// Replace spaces with %20
				DebugMacro.WriteLine("![]("+path+")");
			}
			
			//DebugMacro.WriteLine(knowledgeitem.Address.ToString());
			// DebugMacro.WriteLine(knowledgeitem.Text);
		}
	}
}