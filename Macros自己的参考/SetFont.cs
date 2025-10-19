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

//by qq 1429851250
//插件交流群 202835945

public static class CitaviMacro
{
	public static void Main()
	{
	
		Project project = Program.ActiveProjectShell.Project;

		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		var fontfrm=new FontDialog();
	if(	fontfrm.ShowDialog(mainForm)==DialogResult.OK){
	mainForm.Font=fontfrm.Font;
		MessageBox.Show("设置成功");
	};
		
	}
}