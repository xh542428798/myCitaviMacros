using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.DataExchange;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;

/// <summary>
/// Citavi宏：导入单个文件夹内的PDF并自动分组
/// 
/// 功能概述：
/// 此宏用于将指定单个文件夹中的所有PDF文件批量导入到Citavi项目中。
/// 它会为每个PDF文件创建一个独立的文献条目，并使用所选文件夹的名称
/// 创建一个分组，将所有新导入的文献条目添加到该分组中，实现一键整理。
/// 
/// 工作流程：
/// 1.  提示用户选择一个包含PDF文件的文件夹。
/// 2.  仅扫描该文件夹（不包含其子文件夹）中的全部PDF文件。
/// 3.  对于每一个找到的PDF文件：
///     a. 在Citavi中创建一个新的文献条目，文献类型默认为“Book”。
///     b. 文献标题将使用PDF的文件名（不含扩展名）。
///     c. 将PDF文件本身作为电子附件添加到该文献条目中。
/// 4.  获取用户所选文件夹的名称，查找或创建一个同名分组。
/// 5.  将所有新创建的文献条目统一添加到该分组中。
/// 6.  完成后，弹出一个对话框，显示总共导入的PDF文件数量和分组名称。
/// 
/// 使用场景：
/// 非常适合快速整理单一主题的文档集合，例如：
/// - 将下载的一门课程的全部课件PDF（已放在一个文件夹内）快速导入Citavi，并自动创建以课程名命名的分组。
/// - 将某位作者或某个会议的所有论文PDF一键导入并归类。
/// - 整理任何已经按主题或来源存放在单个文件夹里的PDF文献。
/// 
/// 注意事项：
/// - 运行宏前，Citavi会提示您创建项目备份，以防操作失误。
/// - 此宏不会处理子文件夹中的PDF，仅作用于您选择的当前文件夹。
/// - Citavi的Groups是扁平结构，此宏会创建一个与所选文件夹同名的独立分组。
/// </summary>

public static class CitaviMacro
{
	public static void Main()
	{
	    if (Program.ProjectShells.Count == 0) return;		
	    if (IsBackupAvailable() == false) return;		
		
	    int importedPdfCounter = 0;
		string sourcePath="";
		
	    try 
	    {	
	        
	        FolderBrowserDialog folderDialog = new FolderBrowserDialog();
	        folderDialog.Description = "选择一个包含PDF文件的文件夹...";
	        DialogResult folderResult = folderDialog.ShowDialog();
	        if (folderResult == DialogResult.OK)
	        {
	            sourcePath = folderDialog.SelectedPath;
	        }
	        else return;

	        SwissAcademic.Citavi.Project activeProject = Program.ActiveProjectShell.Project;
	        
	        // 【修改点1】获取选中文件夹的名称，作为统一的分组名
	        string groupName = Path.GetFileName(sourcePath);
	        if (string.IsNullOrEmpty(groupName))
	        {
	            // 如果是根目录，则使用驱动器号作为分组名
	            groupName = sourcePath;
	        }

	        // 【修改点2】只搜索当前文件夹，不包含子文件夹
	        var allPdfFiles = Path2.GetFilesSafe(new DirectoryInfo(sourcePath), "*.pdf", SearchOption.TopDirectoryOnly);

	        foreach (FileInfo pdfFile in allPdfFiles)
	        {
	            string filePath = pdfFile.FullName;
	            string fileName = Path.GetFileNameWithoutExtension(filePath);

	            // 创建一个新的文献条目
	            Reference newReference = new Reference(activeProject);
	            newReference.Title = fileName;
	            newReference.ReferenceType = ReferenceType.Book; // 默认都设为Book

	            // 添加PDF附件
	            Location newlocation = new Location(newReference, LocationType.ElectronicAddress, filePath);
	            newReference.Locations.Add(newlocation);
	            
	            // 【修改点3】使用统一的分组名
	            AddToGroup(newReference, groupName);
	            
	            // 将新文献添加到项目中
	            activeProject.References.Add(newReference);
	            
	            importedPdfCounter++;
	        }
	    } 
	    finally 
	    {
	        string message = string.Format("宏执行完毕。\r\n\r\n已将 {0} 个PDF文件导入，并全部添加到名为 '{1}' 的分组中。", importedPdfCounter, Path.GetFileName(sourcePath));
	        MessageBox.Show(message, "Citavi", MessageBoxButtons.OK, MessageBoxIcon.Information);
	    }
	}

	/// <summary>
	/// 根据给定的分组名称（即文件夹名）为参考文献创建并分配分组。
	/// </summary>
	/// <param name="reference">要处理的参考文献。</param>
	/// <param name="groupName">分组名称。</param>
	private static void AddToGroup(Reference reference, string groupName)
	{
		if (string.IsNullOrEmpty(groupName) || reference == null)
		{
			return;
		}

		Project project = reference.Project;
		
		// 查找或创建分组
		// 注意：Citavi的Groups是扁平的，这里不创建嵌套结构
		var targetGroup = project.Groups.FirstOrDefault(item => item.Name.Equals(groupName, StringComparison.Ordinal));
		if (targetGroup == null)
		{
			// 如果分组不存在，则创建一个新的
			targetGroup = project.Groups.Add(groupName);
		}

		// 将参考文献添加到该分组中
		reference.Groups.Add(targetGroup);
	}

	// IsBackupAvailable 方法保持不变
	private static bool IsBackupAvailable() 
	{
		string warning = String.Concat("重要提示：此宏将对您的项目进行不可逆的更改。",
					"\r\n\r\n", "运行此宏前，请确保您已备份当前项目。",
					"\r\n", "如果您不确定，请点击“取消”，然后在Citavi主窗口的“文件”菜单中选择“创建备份”。",
					"\r\n\r\n", "您是否要继续？"
				);

				
		return (MessageBox.Show(warning, "Citavi", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.OK);
	}
}