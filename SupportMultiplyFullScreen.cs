//运行宏之前，请添加以下引用，引用的文件在bin文件夹下（Citavi安装目录）
// autoref "Infragistics4.Win.UltraWinStatusBar.v11.2.dll"
// autoref "Infragistics4.Win.v11.2.dll"
// autoref "Infragistics4.Win.UltraWinToolbars.v11.2.dll"
// autoref "Infragistics4.Shared.v11.2.dll"
//如果觉得麻烦，可以购买本人已经完成制作的插件，可以默认支持多个Pdf全屏显示。
//作者qq1429851250
//购买插件网址：http://citavi.nat123.net:57484/product.html
//插件售后交流qq群：202835945
//Citavi交流qq群：531215486
//宏编辑日期：2023-10-26 22:40
//宏代码仅用于学习参考，未经本人允许，不得随意出售。
using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using SwissAcademic.Citavi;
using SwissAcademic.Controls;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using SwissAcademic.Citavi.Shell.Controls.Preview;
// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.
using Infragistics.Win;
using Infragistics.Win.UltraWinToolbars;
using Infragistics.Win.UltraWinStatusBar;
public static class CitaviMacro
{
	public static void SetPrivateFieldValue(object obj,string fname,object value){
	obj.GetType().GetField(fname,BindingFlags.NonPublic|BindingFlags.Instance).SetValue(obj, value);
	
	}
	public static T GetValue<T>(object obj,string fname){
	T obj1=(T)obj.GetType().GetField(fname,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(obj);
	return obj1;	
	}
	public static void CreateForm(MainForm mainForm){
	
	MainForm newfrm= (MainForm) Activator.CreateInstance(mainForm.GetType(),true);
newfrm.GetType().BaseType.GetField("_projectShell",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(newfrm, Program.ActiveProjectShell);
		var pctr=newfrm.PreviewControl;
	pctr.Dock=DockStyle.Fill;
 newfrm.Show();
object[] argbool =new object[]{true};
	newfrm.GetType().GetMethod("set_IsPreviewFullScreenForm",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(newfrm,argbool);		
pctr.ShowLocationPreview(mainForm.PreviewControl.ActiveLocation,mainForm.ActiveReference,PreviewBehaviour.Auto,true);		
 PanelWithBorder rightPanelInnerPanel=GetValue<PanelWithBorder>(newfrm,"rightPanelInnerPanel");		
UltraStatusBar statusBar =GetValue<UltraStatusBar>(newfrm,"statusBar");
	newfrm.ActiveRightPaneTab=RightPaneTab.Preview;
Panel p=GetValue<Panel>(newfrm,"mainFormPanel");
ToolbarsManager MainTBM=GetValue<ToolbarsManager>(newfrm,"MainTBM");
		MainTBM.Visible = false;
		p.Visible=false;
		statusBar.Visible=false;
		rightPanelInnerPanel.Controls.Remove(pctr);
		newfrm.Controls.Add(pctr);
		newfrm.Fade();
		newfrm.ColorScheme=mainForm.ColorScheme;
		newfrm.IsCenterPaneCollapsed=true;
		newfrm.Localize();
	    newfrm.LocalizeFormText();	
	}
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm; 

	
		var prevfrm=new Form();
		var bt=new System.Windows.Forms.Button();
		bt.Text="独立打开该引文预览";
		prevfrm.Controls.Add(bt);
		bt.Click += new EventHandler((s,e) => {
		CreateForm(mainForm);
		
		});
		bt.Dock=DockStyle.Fill;
		prevfrm.Show();
		prevfrm.TopMost=true;
	
		
	}

}