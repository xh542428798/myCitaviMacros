// Copyright 2018 István Koren
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
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

		//if this macro should ALWAYS affect all titles in active project, choose first option
		//if this macro should affect just filtered rows if there is a filter applied and ALL if not, choose second option
		
		//ProjectReferenceCollection references = Program.ActiveProjectShell.Project.References;		
		List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		
		//if we need a ref to the active project
		SwissAcademic.Citavi.Project activeProject = Program.ActiveProjectShell.Project;
		
		
		if (references.Count == 2)
		{
			if (references[0].ReferenceType == references[1].ReferenceType)
			{
				string originalTitle = references[0].Title;
				
				// CreatedOn, check which one is older and then take that CreatedOn and CreatedBy
				//if (DateTime.Compare(references[0].CreatedOn, references[1].CreatedOn) > 0)
				//{
					// second reference is older
					// CreatedOn is write-protected. We therefore switch the references...
				//	Reference newer = references[0];
				//	references[0] = references[1];
				//	references[1] = newer;
				//}
				
				// ModifiedOn is write-protected. It will be updated anyways now.
				
				// Title
				if (String.Compare(references[0].Title.Trim(), references[1].Title.Trim(), false) != 0)
				{
					//MessageBox.Show("Title is different:\n"+ references[0].Title + "\n" + references[1].Title);
				}
				
				
				// Abstract, naive approach...
				if (references[0].Abstract.Text.Trim().Length < references[1].Abstract.Text.Trim().Length)
				{
					//MessageBox.Show("Abstract is different:\n"+ references[0].Abstract.Text + "\n" +  references[1].Abstract.Text);
					// references[0].Abstract.Text = references[1].Abstract.Text;
				}
				List<Group> refGroups = references[0].Groups.ToList();
				List<String> myGroupNamelist = new List<string>();
				foreach(Group mygroup in refGroups)
				{
					myGroupNamelist.Add(mygroup.Name);
				}
				myGroupNamelist.Sort();
				MessageBox.Show(string.Join("\n", myGroupNamelist));
				//MessageBox.Show(references[0].Groups.ToList());
				// AccessDate, take newer one
				//TODO: accessdate would need to be parsed
				// right now, we just check if there is one, we take it, otherwise we leave it empty.

				// DONE! remove second reference
				//activeProject.References.Remove(references[1]);

			}
			else
			{
				MessageBox.Show("Currently this script only supports merging two references of the same type. Please convert and try again.");
			}
		}
		else
		{
			MessageBox.Show("Currently this script only supports merging two references. Please select two and try again.");
		}
		
	}
	
	private static string MergeOrCombine(string first, string second) {
		first = first.Trim();
		second = second.Trim();
		
		// do not compare ignore case, otherwise we might lose capitalization information; in that case we rely on manual edits after the merge
		if (String.Compare(first, second, false) == 0)
		{
			// easy case, they are the same!
			return first;
		}
		else if (first.Length == 0)
		{
			return second;
		}
		else if (second.Length == 0)
		{
			return first;
		}
		else
		{
			return first + " // " + second;
		}
	}	
}
