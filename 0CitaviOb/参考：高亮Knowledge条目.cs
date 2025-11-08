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
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Citavi.Shell.Controls.SmartRepeaters;
using System.IO;

// 这个宏只有一部分参考价值，别直接运行

public static class CitaviMacro
{
    public static void Main() 
    {
	    // 1. 
		Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		// 2. 核心：通过遍历控件树来找到 QuotationSmartRepeater，高亮Knowledge的代码
		QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = null;
		if (Program.ActiveProjectShell.PrimaryMainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
        {
            SmartRepeater<KnowledgeItem> KnowledgeItemSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("SmartRepeater", true).FirstOrDefault() as SmartRepeater<KnowledgeItem>;
            quotationSmartRepeaterAsQuotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("knowledgeItemPreviewSmartRepeater", true).FirstOrDefault() as QuotationSmartRepeater;

        }
        else if (Program.ActiveProjectShell.PrimaryMainForm.ActiveWorkspace == MainFormWorkspace.ReferenceEditor)
        {
            quotationSmartRepeaterAsQuotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault() as QuotationSmartRepeater;
        }
		
		if (quotationSmartRepeaterAsQuotationSmartRepeater != null)
		{
		    // 调用方法进行高亮
		    quotationSmartRepeaterAsQuotationSmartRepeater.SelectAndActivate(targetKnowledgeItem, true);
		}
			
		KnowledgeItem knowledge = quotationSmartRepeaterAsQuotationSmartRepeater.ActiveListItem;
		
		
		
		DebugMacro.WriteLine(knowledge.CoreStatement);
		DebugMacro.WriteLine(knowledge.Text);
		DebugMacro.WriteLine(knowledge.PageRange.ToString());
		DebugMacro.WriteLine(knowledge.Id);
        DebugMacro.WriteLine(GetPdfPathFromKnowledgeItem(knowledge));
		
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