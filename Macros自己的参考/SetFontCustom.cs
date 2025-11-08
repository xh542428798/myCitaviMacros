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

//新奥尔良烤乳猪 qq3060191344; 小明 qq1429851250
//citavi交流群 202835945

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