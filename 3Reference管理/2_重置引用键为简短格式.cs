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
		
		List<Reference> references = mainForm.GetSelectedReferences();
		
		foreach (Reference reference in references)
		{
            reference.CitationKey = String.Empty;
			// 获取作者、时间、Title
			Person author = reference.Authors[0];
			// 获取Title并提取前10个单词
			string originalTitle = reference.Title;
			originalTitle = originalTitle.Replace(":", "");
			originalTitle = originalTitle.Replace(",", "");
			string year = reference.Year;
	        string[] words = originalTitle.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
	        string result;
	        if (words.Length >= 3)
	        {// 取前10个单词，首字母大写，其余小写
	            result = string.Join("_", words.Take(3).Select(word => word.First().ToString().ToUpper() + word.Substring(1).ToLower())); 
	        }
	        else
	        {// 所有单词，首字母大写，其余小写
	            result = string.Join("_", words.Select(word => word.First().ToString().ToUpper() + word.Substring(1).ToLower())); 
	        }
			string citationkey = author.LastName.ToString() +year+"_"+result;
			//MessageBox.Show(citationkey);
            if (string.IsNullOrEmpty(reference.CitationKey)) reference.CitationKey =citationkey;// project.CitationKeyAssistant.GenerateKey(reference);
			//reference.CitationKey = project.CitationKeyAssistant.GenerateKey(reference);
		}

	}

	
}