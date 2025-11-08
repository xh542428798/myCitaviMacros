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
		List<KnowledgeItem> knowledges = mainForm.GetSelectedKnowledgeItems();
		foreach (Reference reference in references)
		{
			DebugMacro.WriteLine(reference.Project.Key);
			DebugMacro.WriteLine(reference.Id.ToString());
		}
		foreach (KnowledgeItem knowledge in knowledges)
		{
			DebugMacro.WriteLine(knowledge.FullName);
			DebugMacro.WriteLine(knowledge.Text);
			DebugMacro.WriteLine(knowledge.CoreStatement);
			DebugMacro.WriteLine(knowledge.Id.ToString());
			//await mainForm.PreviewControl.ShowPdfLinkAsync(mainForm, knowledge);
		}
		
	}
}