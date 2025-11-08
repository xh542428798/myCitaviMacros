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
using System.Text.RegularExpressions;
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
		
		int count = 0;
		int totalCount = references.Count;
		foreach (Reference reference in references)
		{
			// your code
			string note = reference.Notes;
			string pattern = @"(\w+):\s(.*?)(?=\s\w+:|$)";
			// 创建正则表达式对象
	        Regex regex = new Regex(pattern);
	        
	        // 在输入字符串中进行匹配
	        MatchCollection matches = regex.Matches(note);

	        // 输出匹配结果
	        foreach (Match match in matches)
	        {
				string title = match.Groups[1].Value;
				string text =  match.Groups[2].Value;
				if(title == "titleTranslation")
				{
					reference.TranslatedTitle = text;
				}
				if(title == "abstractTranslation")
				{
					reference.TableOfContents.Text = text;
				}
				if(title == "Notes")
				{
					reference.TableOfContents.Text = text;
				}
				if(title == "myNote")
				{
					reference.CustomField1 = text;
				}

				if(title == "影响因子")
				{
					reference.CustomField2 = "IF: "+text;
				}
				if(title == "JCR分区")
				{
					reference.CustomField3 = text;
				}
				if(title == "中科院分区升级版")
				{
					reference.CustomField4 = text;
				}
	            //MessageBox.Show(title+": "+text);

	        }
		    count++;
		    DebugMacro.WriteLine(count + " / " + totalCount);
		}
	}
}