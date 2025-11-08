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

/// <summary>
/// Citavi宏：从剪贴板获取PDF路径并在预览中显示
/// 
/// 功能概述：
/// 此宏旨在简化Citavi中查找和预览PDF附件的工作流程。它从Windows剪贴板读取一个PDF文件的路径，
/// 在当前Citavi项目中查找该文件对应的附件，并根据查找结果执行不同的操作。
/// 
/// 工作流程：
/// 1.  获取路径：从剪贴板读取文本，支持本地路径（如 "C:\Docs\file.pdf"）和文件URL（如 "file:///C:/Docs/file.pdf"）。
/// 2.  查找附件：遍历项目中所有的电子附件，将其路径与剪贴板中的路径进行匹配。
/// 3.  执行操作：
///     - 如果找到唯一匹配项：自动切换到该文献条目的附件页面，高亮显示该附件，并在预览区强制刷新并打开此PDF。
///     - 如果未找到任何匹配项：弹出对话框询问用户，确认后以文件名作为标题创建一个新的文献条目，并添加该PDF作为外部附件。
///     - 如果找到多个匹配项（例如，不同文献引用了同名文件）：应用一个临时筛选器，只显示这些匹配的文献，提示用户手动选择。
/// 
/// 技术细节：
/// - 使用反射调用Citavi内部的 `PreviewControl.ShowLocationPreview` 方法，以确保预览区能够可靠地加载并显示目标PDF。
/// - 使用 `Uri.UnescapeDataString` 来正确处理URL编码的路径（如将 %20 转换为空格）。
/// - 通过查找并操作 `LocationSmartRepeater` 控件来精确地在附件列表中定位并激活目标附件。
/// 
/// 使用场景：
/// 当你从文件管理器、PDF阅读器（如Obsidian通过URL Scheme）或其他外部应用复制了一个PDF文件的路径时，
/// 运行此宏可以快速在Citavi项目中定位并预览它，极大提升了文献管理的效率。
/// </summary>

public static class CitaviMacro
{
    /// <summary>
    /// 宏的主入口点，定义为异步方法以支持PDF跳转等异步操作。
    /// </summary>
    public static void Main() 
    {
        // --- 1. 初始化项目与主窗口引用 ---
        // 获取当前活动的Citavi项目和主窗口对象，以便进行后续操作。
        Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
        // --- 2. 初始化用于存储查找结果的列表 ---
        // 获取项目中所有的附件，用于后续查找。
        List<Location> allLocations = mainForm.Project.AllLocations.ToList();
		List<Location> foundLocations = new List<Location>(); // 存储找到的匹配附件
		List<Reference> foundReferences = new List<Reference>(); // 存储找到的匹配附件所属的文献条目
		
        // --- 3. 从剪贴板读取路径并进行预处理 ---
        // 从剪贴板获取文本，假设第一行是文件路径（可能是本地路径或file:// URL）
		string filePath = Clipboard.GetText();
		//DebugMacro.WriteLine(filePath);
		// 直接读取第一行，不会加载整个文件到内存，这通常是文件路径
		string urlPath = File.ReadLines(filePath).FirstOrDefault();
		DebugMacro.WriteLine(urlPath);	
		
		string decodedPath = "";
        // --- 4. 判断路径类型并转换为统一的本地路径格式 ---
        // 检查路径是否以 "file:///" 开头
		if(urlPath.StartsWith("file:///"))
		{
            // 如果是URL格式，则进行解码和转换
			// 去掉开头的 "file:///" (注意是3个斜杠)
			string tempPath = urlPath.Substring(8);
			// 将URL中的 / 替换为 Windows 路径分隔符 \
			string convertedPath = tempPath.Replace('/', '\\');
			// 使用 System.Uri.UnescapeDataString 进行URL解码（处理%20等特殊字符）
			decodedPath = Uri.UnescapeDataString(convertedPath);
		}
		else
		{
            // 如果不是URL格式，则直接使用剪贴板中的原始路径
			decodedPath=filePath;
		}

        // --- 5. 遍历项目中的所有附件，查找匹配项 ---
		int count = 0; // 记录找到的匹配项数量
		Location targetlocation = null; // 用于存储唯一匹配的目标附件
		foreach (Location location in allLocations)
		{
            // 跳过无效地址和非电子附件
			if (string.IsNullOrEmpty(location.Address.ToString())) continue;
			if (location.LocationType != LocationType.ElectronicAddress) continue;
			
            // 获取附件的本地绝对路径
			string locationPath = location.Address.Resolve().LocalPath;
			//DebugMacro.WriteLine(locationPath);

			// 将从剪贴板解析出的路径与当前附件路径进行比较（忽略大小写）
			bool areSame = string.Equals(decodedPath, locationPath, StringComparison.OrdinalIgnoreCase);

			// 4. 使用 string.Format 来组合字符串，而不是 $
			//DebugMacro.WriteLine(string.Format("转换后的路径: {0}", decodedPath));
			//DebugMacro.WriteLine(string.Format("原始本地路径: {0}", locationPath));
			//DebugMacro.WriteLine(string.Format("它们是同一个文件吗? {0}", areSame));
			
			if (areSame)
			{
				//DebugMacro.WriteLine("Yes"); //如果找到匹配项
				count += 1;
				targetlocation = location; // 记录下这个附件
				foundLocations.Add(location); // 添加到找到的附件列表
				foundReferences.Add(location.Reference); // 添加到找到的文献列表
			}
		}
		
        // --- 6. 根据查找结果执行不同操作 ---
		if(count == 1)
		{
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
		else if (count==0)
		{
            // --- 情况二：没有找到匹配项，询问用户是否创建新文献 ---
			// 先弹窗询问用户是否要自动创建
			if (IsCreatRelativeRef() == false) return;
			
			// 从文件路径中提取不含扩展名的文件名作为文献标题
			string fileName = Path.GetFileNameWithoutExtension(decodedPath);

			// 创建一个新的文献条目
			Reference newReference = new Reference(project);
			newReference.Title = fileName;
			newReference.ReferenceType = ReferenceType.JournalArticle; // 默认设置为期刊文章类型

			// 为新文献创建一个电子附件，路径为剪贴板中的路径
			Location newlocation = new Location(newReference, LocationType.ElectronicAddress, decodedPath);
			newReference.Locations.Add(newlocation);
			
			// 将新文献添加到项目中
			project.References.Add(newReference);
		}
		else if (count>1)
		{
            // --- 情况三：找到多个匹配项，应用筛选器让用户选择 ---
			// 创建一个临时筛选器，只显示这些匹配的文献
			ReferenceFilter filter = new ReferenceFilter(foundReferences, "References Selected from Clipboard", false);
            Program.ActiveProjectShell.PrimaryMainForm.ReferenceEditorFilterSet.Filters.ReplaceBy(new List<ReferenceFilter> { filter });
			// 并提示用户
			MessageBox.Show(string.Format("有{0}个的附件同名，已筛选出，请核对",count), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
			Program.ActiveProjectShell.ShowMainForm();
		}
    }
	
	/// <summary>
	/// 弹出一个对话框，询问用户是否要创建一个新的文献条目。
	/// </summary>
	/// <returns>如果用户点击“OK”，则返回 true；否则返回 false。</returns>
	private static bool IsCreatRelativeRef()
	{
		string warning = String.Concat("此文件是真实PDF，在项目中没有被引用，是否自动添加参考文献并以外部路径添加此PDF路径？");
		
		return (MessageBox.Show(warning, "Citavi", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.OK);
	}
}