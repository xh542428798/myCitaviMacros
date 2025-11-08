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

using System.Reflection;


// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{

		
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		//Get the active ("primary") MainForm
		SearchForm searchForm = Program.ActiveProjectShell.SearchForm;
		// 获取类型信息
        Type type = typeof(SearchForm);
		// 获取私有方法信息
        MethodInfo methodInfo = type.GetMethod("GetSelectedReferences", BindingFlags.NonPublic | BindingFlags.Instance);
		
		// 调用私有方法
        
		List<Reference> references= (List<Reference>)methodInfo.Invoke(searchForm, null);
		
		// MessageBox.Show(references[0].Title);
		mainForm.ActiveReference = references[0]; //只会将第1个文献作为Active Reference显示 


		
		
	}
}

