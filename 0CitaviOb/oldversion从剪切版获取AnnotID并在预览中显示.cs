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
using System.Reflection; // 添加反射命名空间，用于调用Citavi内部方法
using SwissAcademic.Citavi.Shell.Controls.Preview;  // 确保命名空间可用，用于访问预览控件
// =================================================================================================
// Citavi 宏：通过剪贴板中的ID定位知识条目并在PDF中跳转
// =================================================================================================
// 
// 功能概述：
// 此宏旨在实现一个高效的工作流：用户在Word或其他外部工具中复制一个或多个Citavi知识条目的ID，
// 然后运行此宏。宏会自动：
// 1. 从剪贴板读取ID。
// 2. 在当前Citavi项目中查找对应的知识条目。
// 3. 将Citavi界面切换到第一个找到的知识条目。
// 4. 如果该知识条目关联了PDF文件，则在PDF预览中精确跳转到该条目对应的高亮或注释位置。
//
// 使用方法：
// 1. 在Citavi中，打开“工具” -> “宏” -> “编辑宏”，将此代码粘贴进去并保存。
// 2. 将一个或多个知识条目的ID（格式为GUID，例如 `09e2e088-f2e6-45d5-a03c-528d82ad64e9`）复制到剪贴板，每个ID占一行。
// 3. 在Citavi中运行此宏。
//
// 注意事项：
// - 宏的实现依赖于Citavi的对象模型，该模型在未来版本中可能会发生变化。
// - 如果知识条目没有关联PDF，或者关联的PDF文件无法访问，宏将弹出提示。
// - 此宏优先处理剪贴板中的第一个有效ID。
//
// =================================================================================================

public static class CitaviMacro
{
    /// <summary>
    /// 宏的主入口点，定义为异步方法以支持PDF跳转等异步操作。
    /// </summary>
    public static async Task Main() 
    {
        // --- 1. 初始化项目与主窗口引用 ---
        // 获取当前活动的Citavi项目和主窗口对象，以便进行后续操作。
        Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
        // --- 2. 从剪贴板读取ID并查找知识条目 ---
        // 获取项目中所有的知识条目，用于后续查找。
        List<Annotation> allAnnotations = project.AllAnnotations.ToList();
        List<Annotation> foundAnnotations = new List<Annotation>();

        // 准备ID列表：用户需要事先将一个或多个知识条目的ID复制到剪贴板，每个ID占一行。
        // 使用 StringReader 逐行读取剪贴板中的文本。
        using (var reader = new StringReader(Clipboard.GetText()))
        {
            for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
            {
                // 在所有知识条目中查找ID与当前行匹配的条目。
                Annotation thisannotation = allAnnotations.Where(k => k.Id.ToString() == line).FirstOrDefault();
                // 如果找不到匹配的条目，则跳过当前行，继续处理下一行。
                if (thisannotation == null) continue;
                // 将找到的知识条目添加到结果列表中。
                foundAnnotations.Add(thisannotation);
            }
        }
		DebugMacro.WriteLine(string.Format("找到Annotation条目数: {0}", foundAnnotations.Count.ToString()));
		//MessageBox.Show(foundKnowledgeItems.Count.ToString());
        // --- 3. 调试代码块 ---
        // 此区域用于开发时的调试输出，在正式使用时可以注释掉。
         foreach (Annotation newannotation in foundAnnotations)
        {
			DebugMacro.WriteLine(string.Format("找到Annotation条目名字: {0}", newannotation.FullName));
			DebugMacro.WriteLine(string.Format("ID: {0}", newannotation.Id.ToString()));
		}

        // --- 4. Citavi界面操作与跳转 ---
        // 检查是否成功找到了至少一个知识条目。
        if(foundAnnotations.Count > 0)
        {
			//return;
            // 4.1. (可选) 在知识组织器中应用筛选
            // 为了让用户看到所有匹配的条目，可以在知识组织器中创建一个临时筛选器。
            // 这一步不是跳转所必需的，但能提供更好的上下文。
            //mainForm.ActiveWorkspace = MainFormWorkspace.KnowledgeOrganizer;
            //KnowledgeItemFilter filter = new KnowledgeItemFilter(foundKnowledgeItems, "Knowledge Items Selected in Word", false);
            //Program.ActiveProjectShell.PrimaryMainForm.KnowledgeOrganizerFilterSet.Filters.ReplaceBy(new List<KnowledgeItemFilter> { filter });
        }
        else
        {
            // 如果剪贴板中没有找到任何有效的ID，则直接退出宏。
            MessageBox.Show("剪贴板中未找到有效的注释条目ID。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
		Annotation annotation = foundAnnotations[0];
		Location targetlocation = annotation.Location;
        // --- 情况一：只找到一个匹配项，直接定位并预览 ---
		// 切换Citavi主界面到文献编辑器，并定位到附件页面
		mainForm.ActiveWorkspace = MainFormWorkspace.ReferenceEditor;
		mainForm.ActiveReferenceEditorTabPage = MainFormReferencesTabPage.TasksAndLocations;
		mainForm.ActiveReference = targetlocation.Reference; // 激活对应的文献条目
		
		// 在附件列表中高亮并激活找到的附件
		Control locationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("locationSmartRepeater", true).FirstOrDefault();
		SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.LocationSmartRepeater locationSmartRepeaterAsLocationSmartRepeater = locationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.LocationSmartRepeater;

		if (locationSmartRepeaterAsLocationSmartRepeater != null)
		{
			locationSmartRepeaterAsLocationSmartRepeater.SelectAndActivate(targetlocation, true);
		}
		
        // --- 强制刷新预览区，显示PDF内容 ---
		try
		{
			var previewControl = mainForm.PreviewControl;
			
			// 使用反射调用Citavi内部的 ShowLocationPreview 方法来加载预览
			Type[] parameterTypes = new Type[] {
				typeof(Location),
				typeof(Reference),
				typeof(SwissAcademic.Citavi.PreviewBehaviour),
				typeof(bool)
			};
			var showLocationPreviewMethod = typeof(PreviewControl).GetMethod("ShowLocationPreview", BindingFlags.Public | BindingFlags.Instance, null, parameterTypes, null);
			
			if (showLocationPreviewMethod != null)
			{
				var previewBehaviourEnum = typeof(SwissAcademic.Citavi.PreviewBehaviour);
				var skipEntryPageValue = Enum.Parse(previewBehaviourEnum, "SkipEntryPage");
				object[] parameters = new object[] { targetlocation, targetlocation.Reference, skipEntryPageValue, true };
				
				showLocationPreviewMethod.Invoke(previewControl, parameters);
				Program.ActiveProjectShell.ShowMainForm();
				//DebugMacro.WriteLine("已调用预览控件渲染PDF。");
			}
			else
			{
				// 如果找不到方法，输出调试信息
				DebugMacro.WriteLine("错误: 找不到匹配的 ShowLocationPreview 方法。");
			}
		}
		catch (Exception ex)
		{
			// 如果调用预览时出错，捕获并记录异常
			DebugMacro.WriteLine("调用预览时发生错误: " + ex.Message);
		}
		
		
		
		
    }
}