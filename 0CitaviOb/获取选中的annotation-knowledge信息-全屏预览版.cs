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

using System.IO;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using System.Reflection;



public static class CitaviMacro
{
    public static void Main() 
    {
       // --- 步骤 1: 获取 ProjectShell 并进行安全检查 ---
        var projectShell = Program.ActiveProjectShell;
        if (projectShell == null)
        {
            MessageBox.Show("错误：当前没有活动的项目窗口。", "宏执行错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- 步骤 2: 通过反射获取 _previewFullScreenForms 字段 ---
        var field = projectShell.GetType().GetField("_previewFullScreenForms", BindingFlags.Instance | BindingFlags.NonPublic);
        if (field == null)
        {
            MessageBox.Show("错误：无法找到 _previewFullScreenForms 字段。", "宏执行错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- 步骤 3: 获取字段的值 ---
        var fullScreenFormsObject = field.GetValue(projectShell);
        if (fullScreenFormsObject == null)
        {
            MessageBox.Show("提示：当前没有打开任何全屏预览窗口。", "调试信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        // --- 关键调试信息 ---
        string objectTypeName = fullScreenFormsObject.GetType().FullName;
        string objectValue = fullScreenFormsObject.ToString();

        string report = "调试信息：\n\n" +"_previewFullScreenForms 字段的实际类型是:\n{objectTypeName}\n\n" +"它的 ToString() 值是:\n{objectValue}";

       //DebugMacro.WriteLine(report);
        // --- 步骤 4: 将对象转换为正确的类型 ---
        var fullScreenForms = fullScreenFormsObject as System.Collections.Generic.List<MainForm>;
        if (fullScreenForms == null)
        {
            MessageBox.Show("错误：无法将全屏窗口列表转换为预期的类型。", "宏执行错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- 步骤 5: 遍历所有全屏窗口并执行你的核心逻辑 ---
        int foundCount = 0;
        foreach (var form in fullScreenForms)
        {
            // 检查这个全屏窗口是否正在预览某个附件
            if (form.PreviewControl != null && form.PreviewControl.ActiveLocation != null)
            {
                var location = form.PreviewControl.ActiveLocation;
                var reference = location.Reference;
                
                // 使用你原来的逻辑，输出到调试窗口
                //DebugMacro.WriteLine(reference.Title);
                foundCount++;
            }
        }

		PreviewControl activePreviewControl = null; // 声明一个变量来存放最终找到的控件

		if (fullScreenForms.Count > 0)
		{
			//DebugMacro.WriteLine("activeFullScreenForm在运行");
		    var activeFullScreenForm = fullScreenForms.LastOrDefault(form => form.Visible);
		    
		    // 在此处设置断点，检查 activeFullScreenForm 是否为 null
		    if (activeFullScreenForm != null)
		    {
		        // 在此处设置断点，确认我们从全屏窗口中找到了一个活动窗口
		        if (activeFullScreenForm.PreviewControl != null)
		        {
		            // 在此处设置断点，确认我们从全屏窗口中找到了控件
		            activePreviewControl = activeFullScreenForm.PreviewControl;
					//DebugMacro.WriteLine("activePreviewControl在运行");
		        }
		    }
		}
		//DebugMacro.WriteLine(string.Format("调试成功！\n\n找到的活动预览控件类型为：{0}", activePreviewControl.GetType().FullName), "调试信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
        // --- 步骤 6: 给用户一个最终的反馈 ---
       	//DebugMacro.WriteLine(string.Format("操作完成。在 {0} 个全屏窗口中，找到了 {1} 个正在预览的文献。", fullScreenForms.Count, foundCount), "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
		
		// 获取PDF预览控件
		object pdfViewControlObject = null;
		// 2. 检查我们获取到的活动预览控件是否有效
		if (activePreviewControl != null)
		{
		    // 3. 从 activePreviewControl 获取 PdfViewControl 属性
		    PropertyInfo pdfViewControlProperty = activePreviewControl.GetType().GetProperty("PdfViewControl", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
		    if (pdfViewControlProperty != null)
		    {
		        // 4. 获取 PdfViewControl 对象
		        pdfViewControlObject = pdfViewControlProperty.GetValue(activePreviewControl);
		    }
		}	

		if (pdfViewControlObject == null)
		{
		    MessageBox.Show("错误：pdfViewControlObject 是 null。");
		    return; // 停止执行
		}
		//DebugMacro.WriteLine("成功：pdfViewControlObject 不为 null。\n类型是: " + pdfViewControlObject.GetType().FullName);

				// 步骤 2: 获取 Tool 对象
		Type pdfViewControlType = pdfViewControlObject.GetType();
		PropertyInfo toolProperty = pdfViewControlType.GetProperty("Tool");

		if (toolProperty == null)
		{
		    MessageBox.Show("错误：找不到 Tool 属性。");
		    return;
		}

		object toolObject = toolProperty.GetValue(pdfViewControlObject);

		if (toolObject == null)
		{
		    MessageBox.Show("错误：Tool 对象是 null。");
		    return;
		}
		// 步骤 3: 获取 Tool 的 SelectedAdornmentContainers 字段
		Type toolType = toolObject.GetType();
		FieldInfo selectedContainersField = toolType.GetField("SelectedAdornmentContainers", BindingFlags.Instance | BindingFlags.NonPublic);

		if (selectedContainersField == null)
		{
		    MessageBox.Show("错误：找不到 SelectedAdornmentContainers 字段。");
		    return;
		}

		object selectedContainersObject = selectedContainersField.GetValue(toolObject);
		if (selectedContainersObject == null)
		{
		    MessageBox.Show("提示：selectedContainersObject 字段是 null。");
		    return;
		}
		// --- 关键修正：我们不再检查它是否为 null，而是直接转换并检查它是否为空 ---
		var selectedContainers = selectedContainersObject as System.Collections.IEnumerable;
		if (selectedContainers == null)
		{
		    MessageBox.Show("错误：无法将 SelectedAdornmentContainers 转换为集合。");
		    return;
		}

		// 使用 .Cast<object>().Any() 来检查集合中是否有任何元素
		// 这比 .GetEnumerator().MoveNext() 更简洁
		if (!selectedContainers.Cast<object>().Any())
		{
		    MessageBox.Show("提示：当前没有选中的高亮或文本。\n\n请尝试用鼠标拖拽选中一段文本，或者单击一个高亮后再试。", "未选中内容", MessageBoxButtons.OK, MessageBoxIcon.Information);
		    return;
		}
		// --- 如果执行到这里，说明 selectedContainers 中有内容，可以安全遍历了 ---
		//DebugMacro.WriteLine(string.Format("成功：找到了 {0} 个选中的装饰容器，准备遍历。", selectedContainers.Cast<object>().Count()));
		
				// 步骤 4: 遍历每个 AdornmentCanvas，获取其 Annotation 属性
		var annotations = new System.Collections.Generic.List<object>();
		foreach (var container in selectedContainers)
		{
			//DebugMacro.WriteLine("var container in selectedContainers在运行");
		    if (container == null) continue;

		    Type containerType = container.GetType();
		    PropertyInfo annotationProperty = containerType.GetProperty("Annotation");
		    if (annotationProperty == null) continue;

		    object annotation = annotationProperty.GetValue(container);
		    if (annotation != null)
		    {
		        annotations.Add(annotation);
		    }
		}
				// 步骤 5: 检查结果
		if (annotations.Count == 0)
		{
		    MessageBox.Show("提示：SelectedAdornmentContainers 集合不为空，但没有找到任何 Annotation 对象。\n请确保您点击的是一个高亮注释，而不是仅仅选中了文本。");
		    return;
		}
		// --- 最终成功 ---
		string message = string.Format("！！！成功获取到 {0} 个 Annotation 对象！！！\n\n", annotations.Count);
		for (int i = 0; i < annotations.Count; i++)
		{
		    object annotation = annotations[i];
		    string typeName = annotation.GetType().FullName;
		    message += string.Format("注释 {0}: 类型为 '{1}'\n", i + 1, typeName);
		}
		//MessageBox.Show(message);
		
		List<KnowledgeItem> knowledges = annotations
		    .OfType<SwissAcademic.Citavi.Annotation>() // 关键步骤：过滤并转换类型
		    .Where(a => a.EntityLinks.Any(e => e.Indication == EntityLink.PdfKnowledgeItemIndication)) // 使用 Any() 更高效
		    .Select(a => (KnowledgeItem)a.EntityLinks.First(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Source)
		    .ToList();

		 // 步骤 1: 尝试获取选中的内容类型
		KnowledgeItem knowledge = knowledges[0];
	
		// 假设你已经获取了这些变量
		string coreStatement = knowledge.CoreStatement;
		string pageRange = knowledge.PageRange.ToString();
		string knowledgeId = knowledge.Id.ToString();
		string pdfPath = GetPdfPathFromKnowledgeItem(knowledge);

		// --- 开始组合链接 ---
		// 1. 处理PDF路径：将反斜杠 \ 替换为正斜杠 /
		string obsidianPath = pdfPath.Replace('\\', '/');
		// 2. 构建链接的显示文本部分
		string linkText = string.Format("{0}/p.{1}", coreStatement, pageRange);
		// 3. 构建完整的Obsidian链接，并在末尾附加上知识ID
		string finalLink = string.Format("[{0}](file:///{1})KnowID：{2}", linkText, obsidianPath, knowledgeId);
		DebugMacro.WriteLine(finalLink);
		Clipboard.SetText(finalLink);
    }

	
    /// <summary>
    /// 核心辅助函数：从一个知识条目对象中，解析出其关联的PDF文件的本地路径。
    /// </summary>
    /// <param name="knowledgeItem">要查询的知识条目。</param>
    /// <returns>如果找到关联的本地PDF文件，则返回其完整路径；否则返回 null。</returns>
    public static string GetPdfPathFromKnowledgeItem(KnowledgeItem knowledgeItem)
    {
        // 步骤 0: 安全检查，确保输入的知识条目不为空。
        if (knowledgeItem == null)
        {
            return null;
        }

        // 步骤 1: 通过 EntityLink 集合查找指向PDF注释的链接。
        // 链接的 Indication 必须是 "PdfKnowledgeItem"，这是Citavi标记PDF引文的标准方式。
        EntityLink pdfLink = null;
        foreach (var el in knowledgeItem.Project.EntityLinks)
        {
            if (el.Source == knowledgeItem && el.Indication == "PdfKnowledgeItem")
            {
                pdfLink = el;
                break; // 找到第一个匹配的链接后立即退出循环。
            }
        }

        // 检查是否找到了有效的链接。
        if (pdfLink == null)
        {
            return null;
        }

        // 检查链接的目标对象是否为 Annotation 类型。
        if (!(pdfLink.Target is SwissAcademic.Citavi.Annotation))
        {
            return null;
        }
        SwissAcademic.Citavi.Annotation pdfAnnotation = (SwissAcademic.Citavi.Annotation)pdfLink.Target;

        // 步骤 2: 从 Annotation 对象获取其所属的 Location 对象。
        var location = pdfAnnotation.Location;
        if (location == null)
        {
            return null;
        }

        // 步骤 3: 从 Location 对象获取其 Address 对象。
        var address = location.Address;
        if (address == null)
        {
            return null;
        }

        // 步骤 4: 调用 Address 的 Resolve() 方法，获取最终的、可访问的 Uri。
        // 这个方法能正确处理本地文件和云附件的缓存路径。
        Uri pdfUri = address.Resolve();

        // 步骤 5: 检查 Uri 是否有效且指向一个本地文件。
        if (pdfUri != null && pdfUri.IsFile)
        {
            // 获取本地文件的完整路径。
            string pdfFilePath = pdfUri.LocalPath;
            return pdfFilePath;
        }
        else
        {
            // 如果不是本地文件（例如是远程URL）或地址无效，则返回 null。
            return null;
        }
    }
	
	
}