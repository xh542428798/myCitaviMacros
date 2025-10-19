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
		searchForm.PerformCommand("BatchApplyGroups");
		// 获取类型信息
        Type type = typeof(SearchForm);
		// 获取私有方法信息
        MethodInfo methodInfo = type.GetMethod("GetSelectedReferences", BindingFlags.NonPublic | BindingFlags.Instance);
		
		// 调用私有方法
        
		List<Reference> references= (List<Reference>)methodInfo.Invoke(searchForm, null);
		
		//List<Reference> list = mainForm.GetFilteredReferences();

		ReferenceFilter referenceFilter = new ReferenceFilter(references, "Search Result", false);
		mainForm.ReferenceEditorFilterSet.Filters.ReplaceBy(new List<ReferenceFilter> { referenceFilter });
		Program.ActivateForm(mainForm);
		

	
		//Program.ActiveProjectShell.ShowApplyGroupsForm(references[0]);
		var commandbarItem = mainForm.GetMainCommandbarManager().GetReferenceEditorCommandbar(MainFormReferenceEditorCommandbarId.Menu).GetCommandbarMenu(MainFormReferenceEditorCommandbarMenuId.BatchModify).GetCommandbarButton("BatchRemoveCategories");
		if (commandbarItem != null)
		{
			
			 // 获取按钮的Click事件处理程序
		    var clickEventHandler = commandbarItem.GetType().GetMethod("Click", BindingFlags.Instance | BindingFlags.NonPublic);

		    // 调用Click事件处理程序以模拟按钮点击
			if (clickEventHandler != null)
			{
			    clickEventHandler.Invoke(commandbarItem, new object[] { EventArgs.Empty });
			}
			

		}
	}
	
	public static void BatchAssignCategories(IList<Reference> references,ProjectShellForm dialogOwner)
    {
      if (references == null || !references.Any<Reference>())
        return;
      if (references.Count == 1)
      {
        dialogOwner.ProjectShell.ShowAssignCategoriesForm(references[0]);
      }
      else
      {
        using (BatchAssignOrRemoveCategoriesDialog categoriesDialog = new BatchAssignOrRemoveCategoriesDialog(dialogOwner, (IEnumerable<Reference>) references, BatchApplyAction.Apply))
        {
          int num = (int) categoriesDialog.ShowDialog((IWin32Window) dialogOwner);
        }
      }
    }
}

