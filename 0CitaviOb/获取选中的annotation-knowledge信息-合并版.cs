using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using System.Reflection;
using System.IO;
using SwissAcademic.Citavi.Shell.Controls.Preview;

public static class CitaviMacro
{
    public static void Main() 
    {
        // --- 核心逻辑：尝试从两种不同的预览控件中获取选中的知识条目 ---
        KnowledgeItem knowledge = TryGetSelectedKnowledgeItem();

        // --- 检查是否成功获取到知识条目 ---
        if (knowledge == null)
        {
            MessageBox.Show("未能获取到选中的知识条目。\n\n请确保：\n1. 你在右侧面板或全屏预览中选中了一个高亮或注释。\n2. 该高亮或注释已经关联到了Citavi的知识条目。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // --- 如果成功，则生成Obsidian链接并复制到剪贴板 ---
        string coreStatement = knowledge.CoreStatement;
        string pageRange = knowledge.PageRange.ToString();
        string knowledgeId = knowledge.Id.ToString();
        string pdfPath = GetPdfPathFromKnowledgeItem(knowledge);

        if (string.IsNullOrEmpty(pdfPath))
        {
            MessageBox.Show("无法获取该知识条目关联的PDF文件路径。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- 开始组合链接 ---
        string obsidianPath = pdfPath.Replace('\\', '/');
        string linkText = string.Format("{0}/p.{1}", coreStatement, pageRange);
        string finalLink = string.Format("[{0}](file:///{1})KnowID：{2}", linkText, obsidianPath, knowledgeId);
        
        DebugMacro.WriteLine("成功生成链接: " + finalLink);
        Clipboard.SetText(finalLink);
        
        MessageBox.Show(string.Format("成功！Obsidian链接已复制到剪贴板。\n{0}",finalLink), "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public static KnowledgeItem TryGetSelectedKnowledgeItem()
    {
        DebugMacro.WriteLine("尝试从右侧面板获取选中的知识条目...");
        PreviewControl mainPreviewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
        KnowledgeItem knowledgeFromMainPanel = GetKnowledgeFromPreviewControl(mainPreviewControl);
        if (knowledgeFromMainPanel != null)
        {
            DebugMacro.WriteLine("成功从右侧面板获取到知识条目。");
            return knowledgeFromMainPanel;
        }
        DebugMacro.WriteLine("右侧面板未找到选中的知识条目。");

        DebugMacro.WriteLine("尝试从全屏预览窗口获取选中的知识条目...");
        KnowledgeItem knowledgeFromFullScreen = GetKnowledgeFromFullScreenPreview();
        if (knowledgeFromFullScreen != null)
        {
            DebugMacro.WriteLine("成功从全屏预览窗口获取到知识条目。");
            return knowledgeFromFullScreen;
        }
        DebugMacro.WriteLine("全屏预览窗口也未找到选中的知识条目。");

        return null;
    }

    public static KnowledgeItem GetKnowledgeFromPreviewControl(PreviewControl previewControl)
    {
        if (previewControl == null) return null;

        PropertyInfo pdfViewControlProperty = previewControl.GetType().GetProperty("PdfViewControl", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
        if (pdfViewControlProperty == null) return null;
        object pdfViewControlObject = pdfViewControlProperty.GetValue(previewControl);
        if (pdfViewControlObject == null) return null;

        PropertyInfo toolProperty = pdfViewControlObject.GetType().GetProperty("Tool");
        if (toolProperty == null) return null;
        object toolObject = toolProperty.GetValue(pdfViewControlObject);
        if (toolObject == null) return null;

        FieldInfo selectedContainersField = toolObject.GetType().GetField("SelectedAdornmentContainers", BindingFlags.Instance | BindingFlags.NonPublic);
        if (selectedContainersField == null) return null;
        object selectedContainersObject = selectedContainersField.GetValue(toolObject);
        if (selectedContainersObject == null) return null;
        
        // 【关键修正】不再直接转换为 IEnumerable，而是通过反射获取枚举器
        IEnumerator enumerator = (selectedContainersObject as IEnumerable).GetEnumerator();
        
        var annotations = new List<object>();
        while (enumerator.MoveNext())
        {
            object container = enumerator.Current;
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

        var knowledges = annotations
            .OfType<SwissAcademic.Citavi.Annotation>()
            .Where(a => a.EntityLinks.Any(e => e.Indication == EntityLink.PdfKnowledgeItemIndication))
            .Select(a => (KnowledgeItem)a.EntityLinks.First(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Source)
            .ToList();

        if (knowledges.Count > 0)
        {
            return knowledges[0];
        }

        return null;
    }

    /// <summary>
    /// 【关键修正】完全使用反射，避免 dynamic 和 foreach
    /// </summary>
    public static KnowledgeItem GetKnowledgeFromFullScreenPreview()
    {
        var projectShell = Program.ActiveProjectShell;
        if (projectShell == null) return null;

        var field = projectShell.GetType().GetField("_previewFullScreenForms", BindingFlags.Instance | BindingFlags.NonPublic);
        if (field == null) return null;
        object fullScreenFormsObject = field.GetValue(projectShell);
        if (fullScreenFormsObject == null) return null;

        // 获取 Count 属性
        PropertyInfo countProperty = fullScreenFormsObject.GetType().GetProperty("Count");
        if (countProperty == null) return null;
        int count = (int)countProperty.GetValue(fullScreenFormsObject);
        if (count == 0) return null;

        // 获取索引器 [int] 以访问元素
        PropertyInfo indexerProperty = fullScreenFormsObject.GetType().GetProperty("Item");
        if (indexerProperty == null) return null;

        MainForm activeFullScreenForm = null;
        // 手动模拟 foreach (var form in fullScreenForms)
        for (int i = 0; i < count; i++)
        {
            object formObject = indexerProperty.GetValue(fullScreenFormsObject, new object[] { i });
            MainForm form = formObject as MainForm;
            if (form != null && form.Visible)
            {
                activeFullScreenForm = form;
            }
        }
        
        if (activeFullScreenForm != null && activeFullScreenForm.PreviewControl != null)
        {
            return GetKnowledgeFromPreviewControl(activeFullScreenForm.PreviewControl);
        }

        return null;
    }

    public static string GetPdfPathFromKnowledgeItem(KnowledgeItem knowledgeItem)
    {
        if (knowledgeItem == null) return null;

        EntityLink pdfLink = null;
        foreach (var el in knowledgeItem.Project.EntityLinks)
        {
            if (el.Source == knowledgeItem && el.Indication == "PdfKnowledgeItem")
            {
                pdfLink = el;
                break;
            }
        }

        if (pdfLink != null && pdfLink.Target is SwissAcademic.Citavi.Annotation)
        {
            SwissAcademic.Citavi.Annotation pdfAnnotation = (SwissAcademic.Citavi.Annotation)pdfLink.Target;
            if (pdfAnnotation != null && pdfAnnotation.Location != null)
            {
                var address = pdfAnnotation.Location.Address;
                if (address != null)
                {
                    Uri pdfUri = address.Resolve();
                    if (pdfUri != null && pdfUri.IsFile)
                    {
                        return pdfUri.LocalPath;
                    }
                }
            }
        }
        return null;
    }
}