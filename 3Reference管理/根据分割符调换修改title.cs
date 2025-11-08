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

		foreach (Reference reference in references)
		{
			string originalTitle = reference.Title.ToString();//reference
			string newTitle = "";
			// 使用 Split 分割字符串，最多分成两部分
        	string[] parts = originalTitle.Split(new[] { '_' }, 2);
			// 如果分割后只有一部分，说明没有找到分隔符，直接返回原字符串
	        if (parts.Length < 2)
	        {
	            newTitle = originalTitle;
	        }
			else
			{
				newTitle = string.Format("{0}_{1}", parts[1], parts[0]);// 重新组合：后半部分 + 分隔符 + 前半部分
				// 将计算出的新标题写回到条目中
				reference.Title = newTitle;
			}

			// 在调试窗口输出结果，方便检查
			DebugMacro.WriteLine(string.Format("原标题: {0} -> 新标题: {1}", originalTitle, newTitle));
		}
	}
}