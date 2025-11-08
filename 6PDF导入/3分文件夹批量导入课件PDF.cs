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
/// Citavi宏：批量导入PDF并按子文件夹自动分组
/// 
/// 功能概述：
/// 此宏用于将指定文件夹及其所有子文件夹中的PDF文件批量导入到Citavi项目中。
/// 它会为每个PDF文件创建一个独立的文献条目，并根据PDF所在的子文件夹名称，
/// 自动将这些文献条目添加到对应的Groups（分组）中，实现自动化整理。
/// 
/// 工作流程：
/// 1.  提示用户选择一个根文件夹。
/// 2.  递归扫描该根文件夹及其所有子文件夹中的全部PDF文件。
/// 3.  对于每一个找到的PDF文件：
///     a. 在Citavi中创建一个新的文献条目，文献类型默认为“Book”。
///     b. 文献标题将使用PDF的文件名（不含扩展名）。
///     c. 将PDF文件本身作为电子附件添加到该文献条目中。
///     d. 获取该PDF所在的直接父文件夹名称，查找或创建一个同名分组，并将此文献条目添加到该分组中。
/// 4.  完成后，弹出一个对话框，显示总共导入的PDF文件数量。
/// 
/// 使用场景：
/// 非常适合整理结构化的电子文档库，例如：
/// - 将按学科、课程或作者分类的课件PDF批量导入Citavi，并自动建立对应的分组。
/// - 将下载的期刊论文集（已按期刊名或年份分好文件夹）一键导入并归类。
/// 
/// 注意事项：
/// - 运行宏前，Citavi会提示您创建项目备份，以防操作失误。
/// - 此宏会创建独立的文献条目，不会建立父子层级关系。
/// - Citavi的Groups是扁平结构，此宏不会创建嵌套的子分组。
/// </summary>
public static class CitaviMacro
{
	public static void Main()
	{
		if (Program.ProjectShells.Count == 0) return;		
		if (IsBackupAvailable() == false) return;		
		
		int importedPdfCounter = 0;
		
		try 
		{	
			string sourcePath;
			FolderBrowserDialog folderDialog = new FolderBrowserDialog();
			folderDialog.Description = "Select a root folder to import all PDFs from it and its subfolders...";
			DialogResult folderResult = folderDialog.ShowDialog();
			if (folderResult == DialogResult.OK)
			{
				sourcePath = folderDialog.SelectedPath;
			}
			else return;

			SwissAcademic.Citavi.Project activeProject = Program.ActiveProjectShell.Project;
			
			// 【核心步骤】递归扫描所有子文件夹中的PDF文件
			var allPdfFiles = Path2.GetFilesSafe(new DirectoryInfo(sourcePath), "*.pdf", SearchOption.AllDirectories);

			foreach (FileInfo pdfFile in allPdfFiles)
			{
				string filePath = pdfFile.FullName;
				string fileName = Path.GetFileNameWithoutExtension(filePath);
				string parentFolderName = Path.GetFileName(Path.GetDirectoryName(filePath)); // 获取PDF文件所在的直接父文件夹名

				// 创建一个新的文献条目
				Reference newReference = new Reference(activeProject);
				newReference.Title = fileName;
				newReference.ReferenceType = ReferenceType.Book; // 默认都设为Book

				// 添加PDF附件
				Location newlocation = new Location(newReference, LocationType.ElectronicAddress, filePath);
				newReference.Locations.Add(newlocation);
				
				// 【关键功能】根据父文件夹名添加到对应的Group
				AddToGroup(newReference, parentFolderName);
				
				// 将新文献添加到项目中
				activeProject.References.Add(newReference);
				
				importedPdfCounter++;
			}
		} 
		finally 
		{
			string message = string.Format("Macro has finished execution.\r\n\r\nImported {0} PDF files as individual references and added them to groups based on their subfolder.", importedPdfCounter);
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