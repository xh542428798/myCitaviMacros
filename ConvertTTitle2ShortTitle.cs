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
			// 将TranslatedTitle转成ShortTitle和Subtitle
			reference.ShortTitle = reference.TranslatedTitle;
			reference.Subtitle = reference.TranslatedTitle;
			// 将Notes转到TableOfValue
			reference.TableOfContents.Text = reference.Notes;
			
			
		}
	}
}