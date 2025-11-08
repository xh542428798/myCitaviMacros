using System;
using System.IO;
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

/* 
详细工作流程
准备ID列表：你事先需要把一个或多个知识条目的ID（例如 123, 456, 789）复制到剪贴板。每个ID占一行。
运行宏：在Citavi中运行这个宏。
宏执行操作：
它会读取剪贴板里的所有ID。
在整个项目中查找这些ID对应的知识条目。
对于找到的每一个条目，它会弹出一个消息框，显示该条目的“核心观点”（Core Statement）。这是一个即时预览功能。
然后，它会根据你当前所在的Citavi界面进行不同的操作：
如果你在“参考文献编辑器”界面：它会为每一个找到的知识条目打开一个独立的编辑窗口。
如果你在“知识组织器”界面：它不会打开新窗口，而是直接应用一个筛选器，只显示这些你指定的知识条目，让你在知识树中集中查看它们。
核心用途
这个宏特别适合以下场景：

从Word文档跳转：正如宏代码中注释的 "Knowledge Items Selected in Word"，它很可能被设计用于与Word插件配合。当你在Word中选中了多个Citavi知识条目并执行某个操作时，Word插件可能会将这些条目的ID复制到剪贴板，然后运行这个宏来在Citavi中定位它们。
批量回顾：你可以手动复制一批相关知识的ID（比如某个章节的所有笔记），然后运行这个宏，快速地在Citavi中逐一查看或集中筛选。
跨软件协作：任何能将文本复制到剪贴板的软件（如Obsidian、Notion等）都可以通过这个宏来与Citavi的知识条目进行交互。 
*/

public static class CitaviMacro
{
	public static void Main()
	{
		Project project = Program.ActiveProjectShell.Project;		
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;		
		
		List<KnowledgeItem> allKnowledgeItems = project.AllKnowledgeItems.ToList();
		List<KnowledgeItem> foundKnowledgeItems = new List<KnowledgeItem>();
		
		using (var reader = new StringReader(Clipboard.GetText()))
		{
		    for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
		    {
		        KnowledgeItem knowledgeItem = allKnowledgeItems.Where(k => k.Id.ToString() == line).FirstOrDefault();
				if (knowledgeItem == null) continue;
				foundKnowledgeItems.Add(knowledgeItem);
				MessageBox.Show(knowledgeItem.CoreStatement);
		    }
		}
		
		if (mainForm.ActiveWorkspace == MainFormWorkspace.ReferenceEditor)
		{
			foreach (KnowledgeItem k in foundKnowledgeItems)
			{				
				Program.ActiveProjectShell.ShowKnowledgeItemFormForExistingItem(mainForm, k);
			}
		}
		else if (mainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
		{
			if(foundKnowledgeItems.Count > 0)
			{
				KnowledgeItemFilter filter = new KnowledgeItemFilter(foundKnowledgeItems, "Knowledge Items Selected in Word", false);
				Program.ActiveProjectShell.PrimaryMainForm.KnowledgeOrganizerFilterSet.Filters.ReplaceBy(new List<KnowledgeItemFilter> { filter });
			}
		}		
	}
}