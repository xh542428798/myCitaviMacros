// autoref "SwissAcademic.Pdf.dll"

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
using System.Diagnostics;
using SwissAcademic.Pdf;
	
// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
	    Project project = Program.ActiveProjectShell.Project;
	    MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

	    // 1. 尝试获取选中的知识条目
	    List<KnowledgeItem> selectedItems = mainForm.GetSelectedQuotations();

	    // 2. 如果没获取到，就尝试另一种方法
	    if (selectedItems == null || selectedItems.Count == 0)
	    {
	        selectedItems = mainForm.GetSelectedKnowledgeItems();
	    }

	    // 3. 【关键修复】检查最终是否有可处理的条目
	    if (selectedItems == null || selectedItems.Count == 0)
	    {
	        MessageBox.Show("请先在Citavi中至少选择一个知识条目（摘录、注释等），然后再运行此宏。", "运行提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
	        return; // 如果两种方法都没获取到东西，就直接退出宏
	    }

	    // 4. 现在可以安全地获取第一个选中的条目了
	    KnowledgeItem knowledgeItem = selectedItems[0];
	    
	    // --- 后续代码保持不变 ---
	    // 检查这个知识条目是否有任何实体链接
	    if (knowledgeItem.EntityLinks != null && knowledgeItem.EntityLinks.Count() > 0)
	    {
	        var targetLink = knowledgeItem.EntityLinks
	            .Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication)
	            .FirstOrDefault();

	        if (targetLink != null && (targetLink.Target is Annotation))
	        {
	            Annotation linkedAnnotation = (Annotation)targetLink.Target;
				
	            PrintAnnotationDetails(linkedAnnotation);
	        }
	    }
	    
	    DebugMacro.WriteLine("运行成功");
		
	}
	
	public static void PrintAnnotationDetails(Annotation annotation)
	{
	    if (annotation == null)
	    {
	        DebugMacro.WriteLine("Annotation is null.");
	        return;
	    }

	    DebugMacro.WriteLine("==================== Annotation Details ====================");
	    DebugMacro.WriteLine(string.Format("ID: {0}", annotation.Id));
	    DebugMacro.WriteLine(string.Format("Visible: {0}", annotation.Visible));

	    // 打印所属的PDF文件信息
	    Location location = annotation.Location;
	    if (location != null)
	    {
	        if (location.Address != null)
	        {
	            DebugMacro.WriteLine(string.Format("PDF File: {0}", location.Address.Resolve().LocalPath));
	        }
	        else
	        {
	            DebugMacro.WriteLine("PDF File: Not found (Address is null).");
	        }
	    }
	    else
	    {
	        DebugMacro.WriteLine("PDF File: Not found (Location is null).");
	    }

	    // 打印颜色信息
	    System.Drawing.Color color = annotation.OriginalColor;
	    Console.WriteLine(string.Format("Color (R,G,B,A): {0}, {1}, {2}, {3}", 
	        color.R, 
	        color.G, 
	        color.B, 
	        color.A));

	    // 打印位置信息
	    if (annotation.Quads != null)
	    {
	        if (annotation.Quads.Count() > 0)
	        {
	            DebugMacro.WriteLine("--- Position & Geometry ---");
			    int segmentIndex = 1;
			    foreach (var quad in annotation.Quads)
			    {
			        DebugMacro.WriteLine(string.Format("  -> Segment {0}:", segmentIndex));
			        DebugMacro.WriteLine(string.Format("     IsContainer: {0}", quad.IsContainer));
			        DebugMacro.WriteLine(string.Format("     PageIndex: {0}", quad.PageIndex));
			        DebugMacro.WriteLine(string.Format("     Bounding Box (X1, Y1, X2, Y2): ({0:F2}, {1:F2}, {2:F2}, {3:F2})", quad.X1, quad.Y1, quad.X2, quad.Y2));
			        segmentIndex++;
			    }
			
	            //foreach (var quad in annotation.Quads)
	            //{
	            //    DebugMacro.WriteLine(string.Format("  -> Segment on Page {0}:", quad.PageIndex));
	            //    DebugMacro.WriteLine(string.Format("     Bottom-Left (X, Y): ({0:F2}, {1:F2})", quad.MinX, quad.MinY));
	            //    DebugMacro.WriteLine(string.Format("     Top-Right (X, Y): ({0:F2}, {1:F2})", quad.MaxX, quad.MaxY));
	            //}
	        }
	        else
	        {
	            DebugMacro.WriteLine("--- Position & Geometry: Quads collection is empty. ---");
	        }
	    }
	    else
	    {
	        DebugMacro.WriteLine("--- Position & Geometry: No Quads found (Quads is null). ---");
	    }

	    DebugMacro.WriteLine("========================================================");
	}

}