// autoref "SwissAcademic.Pdf.dll"

using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using System.IO;
using SwissAcademic.Citavi.Shell.Controls.Preview;

// =================================================================================================
// Citavi 宏：通过剪贴板中的ID定位注释并在PDF中精确跳转 (兼容C# 4.6.1版)
// =================================================================================================

public static class CitaviMacro
{
    /// <summary>
    /// 宏的主入口点。
    /// </summary>
    public static void Main() 
    {
        // --- 1. 初始化项目与主窗口引用 ---
        Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
        // --- 2. 从剪贴板读取ID并查找注释 ---
        List<Annotation> allAnnotations = project.AllAnnotations.ToList();
        List<Annotation> foundAnnotations = new List<Annotation>();

        using (var reader = new StringReader(Clipboard.GetText()))
        {
            for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
            {
                Annotation thisAnnotation = allAnnotations.Where(k => k.Id.ToString() == line.Trim()).FirstOrDefault();
                if (thisAnnotation == null) continue;
                foundAnnotations.Add(thisAnnotation);
            }
        }
		
        if(foundAnnotations.Count == 0)
        {
            MessageBox.Show("剪贴板中未找到有效的注释ID。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // --- 3. 获取第一个找到的注释并准备跳转 ---
        Annotation targetAnnotation = foundAnnotations[0];
        Location targetLocation = targetAnnotation.Location;
        if (targetLocation == null)
        {
            MessageBox.Show("找到的注释没有关联的PDF位置信息。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- 4. 切换Citavi界面到对应文献 ---
        mainForm.ActiveWorkspace = MainFormWorkspace.ReferenceEditor;
        mainForm.ActiveReferenceEditorTabPage = MainFormReferencesTabPage.TasksAndLocations;
        mainForm.ActiveReference = targetLocation.Reference;

        // 在附件列表中高亮并激活找到的附件
        var locationSmartRepeater = mainForm.Controls.Find("locationSmartRepeater", true).FirstOrDefault() as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.LocationSmartRepeater;
        if (locationSmartRepeater != null)
        {
            locationSmartRepeater.SelectAndActivate(targetLocation, true);
        }
		
        // --- 5. 强制刷新预览区，显示PDF内容 ---
        var previewControl = mainForm.PreviewControl;
        try
        {
            Type[] parameterTypes = new Type[] { typeof(Location), typeof(Reference), typeof(SwissAcademic.Citavi.PreviewBehaviour), typeof(bool) };
            var showLocationPreviewMethod = typeof(PreviewControl).GetMethod("ShowLocationPreview", BindingFlags.Public | BindingFlags.Instance, null, parameterTypes, null);
            
            if (showLocationPreviewMethod != null)
            {
                var previewBehaviourEnum = typeof(SwissAcademic.Citavi.PreviewBehaviour);
                var skipEntryPageValue = Enum.Parse(previewBehaviourEnum, "SkipEntryPage");
                object[] parameters = new object[] { targetLocation, targetLocation.Reference, skipEntryPageValue, true };
                
                showLocationPreviewMethod.Invoke(previewControl, parameters);
                Program.ActiveProjectShell.ShowMainForm();
            }
        }
        catch (Exception ex)
        {
            DebugMacro.WriteLine("调用预览时发生错误: " + ex.Message);
            MessageBox.Show("预览PDF时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- 6. 【核心新增部分】执行精确跳转 ---
        // 等待一小段时间，确保PDF完全加载
        System.Threading.Thread.Sleep(1500); 

        // 获取 PdfViewControl 对象
        PropertyInfo pdfViewControlProperty = previewControl.GetType().GetProperty("PdfViewControl", BindingFlags.NonPublic | BindingFlags.Instance);
        object pdfViewControlObject = null;
        if (pdfViewControlProperty != null)
        {
            pdfViewControlObject = pdfViewControlProperty.GetValue(previewControl);
        }

        if (pdfViewControlObject != null)
        {
            // --- 【最终修改后的代码】 ---
            // 1. 直接使用 GetMethod 查找有两个参数的 GoToAnnotation 方法
            Type[] parameterTypes = new Type[] { typeof(Annotation), typeof(EntityLink) };
            MethodInfo goToAnnotationMethod = pdfViewControlObject.GetType().GetMethod("GoToAnnotation", parameterTypes);

            // 2. 如果找到了方法，就调用它
            if (goToAnnotationMethod != null)
            {
                try
                {
                    // 调用时传入两个参数：目标注释和 null
                    object result = goToAnnotationMethod.Invoke(pdfViewControlObject, new object[] { targetAnnotation, null });
                    
                    if (result is bool && !(bool)result)
                    {
                        MessageBox.Show("PDF已加载，但未能跳转到指定注释。可能该注释在当前PDF中不可见。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("调用 GoToAnnotation 时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                // 如果还是找不到，打印出所有带参数的方法进行终极调试
                DebugMacro.WriteLine("错误：在PdfViewControl中未找到 GoToAnnotation(Annotation, EntityLink) 方法。");
                DebugMacro.WriteLine("PdfViewControl 所有带参数的方法列表：");
                foreach (var method in pdfViewControlObject.GetType().GetMethods())
                {
                    var paramInfos = method.GetParameters();
                    if (paramInfos.Length > 0)
                    {
                        string paramList = string.Join(", ", paramInfos.Select(p => p.ParameterType.Name + " " + p.Name));
                        DebugMacro.WriteLine("- " + method.Name + "(" + paramList + ")");
                    }
                }
                MessageBox.Show("在PdfViewControl中未找到 GoToAnnotation 方法。请查看调试窗口获取更多信息。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        else
        {
            MessageBox.Show("无法获取PdfViewControl对象。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


    }
}