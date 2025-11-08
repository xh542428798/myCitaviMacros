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
/*
 * =====================================================================================
 * 宏名称：层级化PDF导入工具 (Groups版)
 * 版本：V2.0
 * 作者：[你的名字]
 * 创建日期：2023-10-27
 * 
 * -------------------------------------------------------------------------------------
 * 功能描述：
 * 本宏用于将按文件夹组织的拆分PDF文件批量导入Citavi，并自动创建规范的层级结构。
 * 它特别适用于将一本电子书的各个章节整理成一个“编辑书籍”及其下的多个“贡献”。
 * 与V1.0不同，此版本使用Groups（分组）而非Categories（类别）进行组织。
 *
 * -------------------------------------------------------------------------------------
 * 前置条件：
 * 1. 请确保你的文件结构是：根文件夹 -> 多个一级子文件夹 -> 每个子文件夹内包含多个PDF文件。
 *    示例：
 *    E:\MyBooks\
 *    ├── 格氏解剖学\
 *    │   ├── 01_绪论.pdf
 *    │   └── 02_上肢.pdf
 *    └── 奈特人体解剖学图谱\
 *        ├── 01_头部.pdf
 *        └── 02_颈部.pdf
 *
 * -------------------------------------------------------------------------------------
 * 执行流程：
 * 1. 【阶段一：创建父条目】
 *    - 扫描根目录下的所有一级子文件夹。
 *    - 为每个子文件夹创建一个类型为 "BookEdited" (编辑书籍) 的父条目。
 *    - 父条目的标题设置为子文件夹的名称。
 *    - 自动为父条目添加一个同名的Group（分组）。
 *
 * 2. 【阶段二：创建子条目】
 *    - 遍历每个父条目对应的文件夹。
 *    - 为文件夹内的每个PDF文件创建一个类型为 "Contribution" (贡献) 的子条目。
 *    - 子条目的标题设置为PDF文件名（不含扩展名）。
 *    - 自动将子条目的 ParentReference 属性链接到其父条目，建立层级关系。
 *    - 为子条目添加与父条目相同的Group（分组）。
 *    - 为子条目创建一个指向原始PDF文件的链接，不复制文件。
 *
 * 3. 【阶段三：处理独立PDF】
 *    - 扫描根目录下的PDF文件。
 *    - 为每个PDF创建一个独立的 "Book" 类型条目，不添加任何Group。
 *
 * -------------------------------------------------------------------------------------
 * 最终效果：
 * 在Citavi项目中，你会得到一个清晰的层级结构和分组：
 * - 父条目 (BookEdited): "格氏解剖学" (位于 "格氏解剖学" Group)
 *   ├─ 子条目 (Contribution): "01_绪论" (位于 "格氏解剖学" Group, 附件: 链接)
 *   └─ 子条目 (Contribution): "02_上肢" (位于 "格氏解剖学" Group, 附件: 链接)
 *
 * =====================================================================================
 */
public static class CitaviMacro
{
	public static void Main()
	{
		if (Program.ProjectShells.Count == 0) return;		
		if (IsBackupAvailable() == false) return;		
		
		int parentCounter = 0;
		int childCounter = 0;
		int singleBookCounter = 0;
		
		try 
		{	
			string sourcePath;
			FolderBrowserDialog folderDialog = new FolderBrowserDialog();
			folderDialog.Description = "Select a root folder for the books (each subfolder is a book)...";
			DialogResult folderResult = folderDialog.ShowDialog();
			if (folderResult == DialogResult.OK)
			{
				sourcePath = folderDialog.SelectedPath;
			}
			else return;

			SwissAcademic.Citavi.Project activeProject = Program.ActiveProjectShell.Project;
			
			// 【步骤1】扫描所有一级子文件夹，并创建父条目
			var subDirectories = Directory.GetDirectories(sourcePath);
			Dictionary<string, Reference> parentReferences = new Dictionary<string, Reference>(StringComparer.OrdinalIgnoreCase);

			foreach (string subDirPath in subDirectories)
			{
				string parentFolderName = Path.GetFileName(subDirPath);
				Reference parentReference = new Reference(activeProject);
				parentReference.Title = parentFolderName;
				parentReference.ReferenceType = ReferenceType.BookEdited;
				activeProject.References.Add(parentReference);

				// 【修改点】为父条目添加到同名的Group
				AddToGroup(parentReference, parentFolderName);

				parentReferences[subDirPath] = parentReference;
				parentCounter++;
			}

			// 【步骤2】扫描所有PDF文件，创建子条目并关联父条目
			foreach (var parentPair in parentReferences)
			{
				string parentDirPath = parentPair.Key;
				Reference parentReference = parentPair.Value;

				var pdfFiles = Path2.GetFilesSafe(new DirectoryInfo(parentDirPath), "*.pdf", SearchOption.TopDirectoryOnly);

				foreach (FileInfo pdfFile in pdfFiles)
				{
					string filePath = pdfFile.FullName;
					string directoryPath = Path.GetDirectoryName(filePath);
					string parentFolderName = Path.GetFileName(directoryPath);
					string thisfileName = Path.GetFileNameWithoutExtension(filePath);
					string fileName = string.Format("{0}_{1}", parentFolderName, thisfileName);

					Reference childReference = new Reference(activeProject);
					childReference.Title = fileName;
					childReference.ReferenceType = ReferenceType.Contribution;
					childReference.ParentReference = parentReference;

					Location newlocation = new Location(childReference, LocationType.ElectronicAddress, filePath);
					childReference.Locations.Add(newlocation);
					
					activeProject.References.Add(childReference);

					// 【修改点】为子条目也添加到和父条目相同的Group
					AddToGroup(childReference, Path.GetFileName(parentDirPath));
					
					childCounter++;
				}
			}
			
			// 【步骤3】处理根目录下的所有独立PDF（创建单本书）
			var rootPdfFiles = Path2.GetFilesSafe(new DirectoryInfo(sourcePath), "*.pdf", SearchOption.TopDirectoryOnly);

			foreach (FileInfo pdfFile in rootPdfFiles)
			{
				string filePath = pdfFile.FullName;
				string fileName = Path.GetFileNameWithoutExtension(filePath);

				Reference newReference = new Reference(activeProject);
				newReference.Title = fileName;
				newReference.ReferenceType = ReferenceType.Book;

				Location newlocation = new Location(newReference, LocationType.ElectronicAddress, filePath);
				newReference.Locations.Add(newlocation);
				
				activeProject.References.Add(newReference);
				
				singleBookCounter++;
			}
		} 
		finally 
		{
			string message = string.Format("Macro has finished execution.\r\n\r\nCreated:\r\n- {0} parent books (BookEdited)\r\n- {1} child chapters (Contribution)\r\n- {2} single books (Book)", parentCounter, childCounter, singleBookCounter);
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
