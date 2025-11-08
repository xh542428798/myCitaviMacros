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
	    // 1. 
		Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
	   // 获取PDF预览控件
		object pdfViewControlObject = null;
		if (Program.ActiveProjectShell.PrimaryMainForm != null)
		{
		    if (Program.ActiveProjectShell.PrimaryMainForm.PreviewControl != null)
		    {
		        PropertyInfo pdfViewControlProperty = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl.GetType().GetProperty("PdfViewControl", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
		        if (pdfViewControlProperty != null)
		        {
		            pdfViewControlObject = pdfViewControlProperty.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl);
		        }
		    }
		}		


		// 步骤 2: 获取 Tool 对象
		Type pdfViewControlType = pdfViewControlObject.GetType();
		PropertyInfo toolProperty = pdfViewControlType.GetProperty("Tool");


		object toolObject = toolProperty.GetValue(pdfViewControlObject);

		//MessageBox.Show("成功：获取到 Tool 对象。\n类型是: " + toolObject.GetType().FullName);

		// 步骤 3: 获取 Tool 的 SelectedAdornmentContainers 字段 (这是核心！)
		Type toolType = toolObject.GetType();
		FieldInfo selectedContainersField = toolType.GetField("SelectedAdornmentContainers", BindingFlags.Instance | BindingFlags.NonPublic);

		object selectedContainersObject = selectedContainersField.GetValue(toolObject);

		if (selectedContainersObject == null)
		{
		    //MessageBox.Show("提示：SelectedAdornmentContainers 字段是 null。");
		    return;
		}
		// 将结果转换为 IEnumerable 以便遍历
		var selectedContainers = selectedContainersObject as System.Collections.IEnumerable;
		// 步骤 4: 遍历每个 AdornmentCanvas，获取其 Annotation 属性
		var annotations = new System.Collections.Generic.List<object>();
		foreach (var container in selectedContainers)
		{
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
		    //MessageBox.Show("提示：SelectedAdornmentContainers 集合不为空，但没有找到任何 Annotation 对象。\n请确保您点击的是一个高亮注释，而不是仅仅选中了文本。");
		    return;
		}
		// --- 最终成功 ---
		string message = string.Format("！！！成功获取到 {0} 个 Annotation 对象！！！\n\n", annotations.Count);
		for (int i = 0; i < annotations.Count; i++)
		{
		    object annotation = annotations[i];
		    string typeName = annotation.GetType().FullName;
		    //message += string.Format("注释 {0}: 类型为 '{1}'\n", i + 1, typeName);
		}
		//MessageBox.Show(message);


		List<KnowledgeItem> knowledges = annotations
		    .OfType<SwissAcademic.Citavi.Annotation>() // 关键步骤：过滤并转换类型
		    .Where(a => a.EntityLinks.Any(e => e.Indication == EntityLink.PdfKnowledgeItemIndication)) // 使用 Any() 更高效
		    .Select(a => (KnowledgeItem)a.EntityLinks.First(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Source)
		    .ToList();

		 // 步骤 1: 尝试获取选中的内容类型
		KnowledgeItem knowledge = knowledges[0];


		//DebugMacro.WriteLine(knowledge.CoreStatement);
		//DebugMacro.WriteLine(knowledge.Text);
		//DebugMacro.WriteLine(knowledge.PageRange.ToString());
		//DebugMacro.WriteLine(knowledge.Id);
        //DebugMacro.WriteLine(GetPdfPathFromKnowledgeItem(knowledge));
		
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