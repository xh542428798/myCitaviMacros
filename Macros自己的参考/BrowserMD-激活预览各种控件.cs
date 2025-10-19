using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;  
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using Markdig;
using System.Reflection;
using SwissAcademic.Citavi.Shell.Controls.Preview;  // 确保命名空间可用
using System.Net;
// autoref "E:\Software\CitaviTest\Addons\Markdig.dll"
// autoref "E:\Software\CitaviTest\bin\System.IO.dll"
// autoref "E:\Software\CitaviTest\bin\System.Runtime.dll"

// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

		//if this macro should affect just filtered rows in the active MainForm, choose:
		List<Reference> references = mainForm.GetSelectedReferences();	
        if (references == null || references.Count == 0)
        {
            MessageBox.Show("请先在文献列表中选中一个包含 .md 附件的文献条目。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        Reference reference = references[0];
        if (reference.Locations == null || !reference.Locations.Any())
        {
            MessageBox.Show("选中的文献条目没有任何附件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 查找第一个符合条件的MD文件Location
        Location location = reference.Locations.FirstOrDefault(loc =>
            loc.Address.LinkedResourceType == LinkedResourceType.AbsoluteFileUri ||
            loc.Address.LinkedResourceType == LinkedResourceType.RelativeFileUri);

        if (location == null)
        {
            MessageBox.Show("选中的文献条目没有本地文件附件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        Uri uri = location.Address.Resolve();
        if (uri == null || !uri.IsFile || !uri.LocalPath.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("选中的文献条目没有找到本地 Markdown (.md) 文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 异步执行预览任务，避免阻塞UI
        PreviewMarkdownAsync(location, mainForm);
    }

  // 异步执行预览的核心逻辑
    private static async void PreviewMarkdownAsync(Location location, MainForm mainForm)
    {
        try
        {
            string mdFilePath = location.Address.Resolve().GetLocalPathSafe();

            // 1. 将MD转换为HTML字符串
            string markdown = File.ReadAllText(mdFilePath);
            string htmlBody = Markdown.ToHtml(markdown, new MarkdownPipelineBuilder().UseAdvancedExtensions().Build());
            DebugMacro.WriteLine("MD转换成功。");
			//DebugMacro.WriteLine(htmlBody);

            // 2. 构建完整的HTML文档
            string fullHtml = string.Format(@"
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset='utf-8'>
                    <title>{0}</title>
                    <style>
                        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif; margin: 2em; line-height: 1.6; color: #333; }}
                        h1, h2, h3, h4, h5, h6 {{ color: #1a1a1a; margin-top: 1.5em; margin-bottom: 0.5em; }}
                        p {{ margin-bottom: 1em; }}
                        a {{ color: #0066cc; }}
                        code {{ background-color: #f4f4f4; padding: 0.2em 0.4em; border-radius: 3px; font-size: 0.9em; }}
                        pre {{ background-color: #f4f4f4; padding: 1em; border-radius: 5px; overflow-x: auto; white-space: pre-wrap; }}
                        blockquote {{ border-left: 4px solid #ddd; margin: 0; padding-left: 1em; color: #666; }}
                        table {{ border-collapse: collapse; width: 100%; margin-bottom: 1em; }}
                        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                        th {{ background-color: #f2f2f2; }}
                    </style>
                </head>
                <body>
                    {1}
                </body>
                </html>", location.Reference.ShortTitle, htmlBody);

            // 3. 获取预览控件并执行渲染
            var previewControl = mainForm.PreviewControl;
            RenderMarkdownInPreviewControl(previewControl, fullHtml, location);

            DebugMacro.WriteLine("已调用预览控件渲染HTML。");
        }
        catch (Exception ex)
        {
            DebugMacro.WriteLine(string.Format("预览过程中发生错误: {0}\n{1}", ex.Message, ex.StackTrace));
            MessageBox.Show(string.Format("预览失败: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

// *** 最终融合方案：先激活内容，再绘制Epub界面 ***
private static void RenderMarkdownInPreviewControl(PreviewControl previewControl, string fullHtml, Location location)
{
    DebugMacro.WriteLine("--- 开始执行最终融合版 RenderMarkdownInPreviewControl ---");

    // --- 第一部分：使用 ShowLocationPreview 激活浏览器并加载内容 ---
    DebugMacro.WriteLine("--- 第一部分：激活浏览器并加载内容 ---");
    string tempHtmlPath = "";
    try
    {
        // 1. 创建一个临时的HTML文件
        tempHtmlPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "citavi_temp_" + Guid.NewGuid().ToString("N") + ".html");
        System.IO.File.WriteAllText(tempHtmlPath, fullHtml);
        DebugMacro.WriteLine("步骤 1.1: 临时HTML文件已创建。");

        // 2. 创建一个临时的Location对象
        var tempLocation = new Location(location.Reference, LocationType.ElectronicAddress, tempHtmlPath);
        DebugMacro.WriteLine("步骤 1.2: 临时Location对象已创建。");

        // 3. 精确匹配并调用 ShowLocationPreview
        Type[] parameterTypes = new Type[] {
            typeof(Location),
            typeof(Reference),
            typeof(SwissAcademic.Citavi.PreviewBehaviour),
            typeof(bool)
        };
        var showLocationPreviewMethod = typeof(PreviewControl).GetMethod("ShowLocationPreview", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, parameterTypes, null);
        
        if (showLocationPreviewMethod != null)
        {
            var previewBehaviourEnum = typeof(SwissAcademic.Citavi.PreviewBehaviour);
            var skipEntryPageValue = Enum.Parse(previewBehaviourEnum, "SkipEntryPage");
            object[] parameters = new object[] { tempLocation, location.Reference, skipEntryPageValue, true };
            
            showLocationPreviewMethod.Invoke(previewControl, parameters);
            DebugMacro.WriteLine("步骤 1.3: ShowLocationPreview() 调用成功，浏览器应已激活并加载内容。");
        }
        else
        {
            DebugMacro.WriteLine("错误: 找不到匹配的 ShowLocationPreview 方法。");
        }

        // 4. 等待Citavi处理
        System.Threading.Tasks.Task.Delay(200).Wait();
    }
    catch (Exception ex)
    {
        DebugMacro.WriteLine("在第一部分发生错误: " + ex.Message);
    }
    finally
    {
        // 5. 清理临时文件
        if (!string.IsNullOrEmpty(tempHtmlPath) && System.IO.File.Exists(tempHtmlPath))
        {
            try { System.IO.File.Delete(tempHtmlPath); }
            catch { /* 忽略删除错误 */ }
        }
    }

    // --- 第二部分：使用反射强制绘制Epub界面 ---
    DebugMacro.WriteLine("--- 第二部分：强制绘制Epub界面 ---");
    try
    {
        // 6. (关键步骤) 设置内部状态字段为 Epub
        var activePreviewTypeField = typeof(PreviewControl).GetField("_activePreviewType", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        if (activePreviewTypeField != null)
        {
            activePreviewTypeField.SetValue(previewControl, PreviewType.Epub);
            DebugMacro.WriteLine("步骤 2.1: _activePreviewType 已强制设置为 Epub。");
        }

        // 7. 调用私有的、更底层的更新方法（这是调出Bookmarks界面的关键）
        var updateToolbarForPreviewTypeMethod = typeof(PreviewControl).GetMethod("UpdateToolbarForPreviewType", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        if(updateToolbarForPreviewTypeMethod != null)
        {
            updateToolbarForPreviewTypeMethod.Invoke(previewControl, null);
            DebugMacro.WriteLine("步骤 2.2: UpdateToolbarForPreviewType() 已调用。");
        }

        var updateStatusBarForPreviewTypeMethod = typeof(PreviewControl).GetMethod("UpdateStatusBarForPreviewType", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        if(updateStatusBarForPreviewTypeMethod != null)
        {
            updateStatusBarForPreviewTypeMethod.Invoke(previewControl, null);
            DebugMacro.WriteLine("步骤 2.3: UpdateStatusBarForPreviewType() 已调用。");
        }
    }
    catch (Exception ex)
    {
        DebugMacro.WriteLine("在第二部分发生错误: " + ex.Message);
    }

    // --- 第三部分：直接修改SmartRepeater的内部选中项 ---
    DebugMacro.WriteLine("--- 第三部分：直接修改SmartRepeater的内部选中项 ---");
    try
    {
        var mainForm = previewControl.TopLevelControl as MainForm;
        if (mainForm != null)
        {
            var smartRepeaterField = typeof(MainForm).GetField("uriLocationsSmartRepeater", BindingFlags.NonPublic | BindingFlags.Instance);
            if (smartRepeaterField != null)
            {
                var smartRepeater = smartRepeaterField.GetValue(mainForm);
                if (smartRepeater != null)
                {
                    // 尝试找到存储当前选中项的私有字段，通常叫 _activeListItem 或类似的名字
                    var activeListItemField = smartRepeater.GetType().GetField("_activeListItem", BindingFlags.NonPublic | BindingFlags.Instance) ??
                                            smartRepeater.GetType().GetField("activeListItem", BindingFlags.NonPublic | BindingFlags.Instance);

                    if (activeListItemField != null)
                    {
                        // 直接将我们原始的Location对象设置给它
                        activeListItemField.SetValue(smartRepeater, location);
                        DebugMacro.WriteLine("步骤 3.1: 已成功设置 _activeListItem 字段，UI选择状态已恢复。");
                    }
                    else
                    {
                        DebugMacro.WriteLine("错误: 在 SmartRepeater 中找不到 _activeListItem 字段。");
                    }
                }
            }
        }
    }
    catch (Exception ex)
    {
        DebugMacro.WriteLine("在第三部分发生错误: " + ex.Message);
    }
	
    DebugMacro.WriteLine("--- 最终融合版 RenderMarkdownInPreviewControl 执行完毕 ---");
}



}