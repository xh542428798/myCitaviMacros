// 将项目里的全部外部链接批量转换到citavi默认附件文件夹下，注意33行
//将外部链接的附件（如PDF）批量移动或复制到Citavi项目的默认附件文件夹中。

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


public static class CitaviMacro
{
	public static void Main()
	{
		bool deleteLocationAfterResolve = true; // True to delete the folder location after it has been resolved
		AttachmentAction attachmentAction = AttachmentAction.Move; // Set the attachment action for the files in the directory. Note that for cloud and sql server project this will always be resetted to Copy.

		int counter = 0;
		
		SwissAcademic.Citavi.Project activeProject = Program.ActiveProjectShell.Project;
		DebugMacro.WriteLine(activeProject.Addresses.AttachmentsFolderPath);//简单的查看项目的默认附件文件夹
		string projectCitaviFilesFolderPath = activeProject.Addresses.AttachmentsFolderPath;
		
		try	
		{
			foreach(Location location in Program.ActiveProjectShell.Project.AllLocations.ToList()) //注意这里行
			{
				//DebugMacro.WriteLine(location.Address.Resolve().ToString());
				if (string.IsNullOrEmpty(location.Address.ToString())) continue;
				if (location.LocationType != LocationType.ElectronicAddress) continue;
				string filePath = location.Address.ToString();
				//DebugMacro.WriteLine(location.Address.ToString());
				//DebugMacro.WriteLine(location.LocationType.ToString());
				if(filePath.Contains("OneDrive - xiehui1573"))
				{
					counter++;
					location.Reference.Locations.Add(filePath, attachmentAction, AttachmentNaming.Rename);
					//location.ResolveFolderLocation(attachmentAction);//这里是处理文件夹的，不是文件
					
					if(deleteLocationAfterResolve)
					{
						location.Reference.Locations.Remove(location);
					}
				}
			}
			MessageBox.Show("Resolved " + counter.ToString() + " locations");
		}
		catch(Exception x)
		{
			MessageBox.Show(x.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
	
	}
}