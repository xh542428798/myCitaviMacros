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
using System.IO;
// =================================================================================================
// Citavi 宏：通过剪贴板中的ID定位文献条目并在预览中显示
// =================================================================================================
// 
// 功能概述：
// 此宏旨在实现一个高效的工作流：用户在Word或其他外部工具中复制一个或多个Citavi文献条目的ID，
// 然后运行此宏。宏会自动：
// 1. 从剪贴板读取ID。
// 2. 在当前Citavi项目中查找对应的文献条目。
// 3. 将Citavi界面切换到第一个找到的文献条目。
// 4. 如果该文献条目有关联的附件，则在预览区中显示该附件。
//
// 使用方法：
// 1. 在Citavi中，打开“工具” -> “宏” -> “编辑宏”，将此代码粘贴进去并保存。
// 2. 将一个或多个文献条目的ID（格式为GUID，例如 `09e2e088-f2e6-45d5-a03c-528d82ad64e9`）复制到剪贴板，每个ID占一行。
// 3. 在Citavi中运行此宏。
//
// 注意事项：
// - 宏的实现依赖于Citavi的对象模型，该模型在未来版本中可能会发生变化。
// - 如果文献条目没有关联任何附件，或者关联的附件无法访问，预览区将显示为空。
// - 此宏会优先处理剪贴板中的第一个有效ID，并将其设为当前活动文献。
// - 如果剪贴板中有多个ID，宏会应用一个临时筛选器，以显示所有匹配的文献条目。
//
// =================================================================================================

public static class CitaviMacro
{

    public static void Main() 
    {
        // --- 1. 初始化项目与主窗口引用 ---
        // 获取当前活动的Citavi项目和主窗口对象，以便进行后续操作。
        Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
        // --- 2. 从剪贴板读取ID并查找文献条目 ---
        // 获取项目中所有的文献条目，用于后续查找。
        List<Reference> allReferences = project.References.ToList();
        List<Reference> foundReferences = new List<Reference>();

        // 准备ID列表：用户需要事先将一个或多个文献条目的ID复制到剪贴板，每个ID占一行。
        // 使用 StringReader 逐行读取剪贴板中的文本。
        using (var reader = new StringReader(Clipboard.GetText()))
        {
            for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
            {
                // 在所有文献条目中查找ID与当前行匹配的条目。
                Reference reference = allReferences.Where(k => k.Id.ToString() == line).FirstOrDefault();
                // 如果找不到匹配的条目，则跳过当前行，继续处理下一行。
                if (reference == null) continue;
                // 将找到的文献条目添加到结果列表中。
                foundReferences.Add(reference);
            }
        }
		
        // --- 3. 调试代码块 ---
        // 此区域用于开发时的调试输出，在正式使用时可以注释掉。
        // foreach (KnowledgeItem knowledge in foundKnowledgeItems)
        // {
        //     DebugMacro.WriteLine("找到知识条目: {knowledge.FullName}");
        //     DebugMacro.WriteLine("核心思想: {knowledge.CoreStatement}");
        //     DebugMacro.WriteLine("ID: {knowledge.Id.ToString()}");
        // }

        // --- 4. 检查查找结果并执行界面操作 ---
        // 检查是否成功找到了至少一个文献条目。
        if(foundReferences.Count > 0)
        {
            // 3.1. 将第一个找到的文献条目设为当前活动文献
            // 这会使得主窗体右侧的预览区立即显示该文献的详情或附件。
			mainForm.ActiveWorkspace = MainFormWorkspace.ReferenceEditor;
            Reference targetReference = foundReferences.FirstOrDefault();
            mainForm.ActiveReference = targetReference;
			Program.ActiveProjectShell.ShowMainForm();
            // 3.2. (可选) 在文献编辑器中应用筛选
            // 如果剪贴板中有多个ID，为了在文献列表中高亮显示所有匹配的条目，
            // 可以在文献编辑器中创建一个临时筛选器。
            if (foundReferences.Count > 1)
            {
                ReferenceFilter filter = new ReferenceFilter(foundReferences, "References Selected from Clipboard", false);
                Program.ActiveProjectShell.PrimaryMainForm.ReferenceEditorFilterSet.Filters.ReplaceBy(new List<ReferenceFilter> { filter });
				Program.ActiveProjectShell.ShowMainForm();
            }
        }
        else
        {
            // 如果剪贴板中没有找到任何有效的ID，则提示用户。
            MessageBox.Show("剪贴板中未找到有效的文献条目ID。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
    }
}