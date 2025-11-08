// autoref "E:\Software\CitaviTest\bin\SwissAcademic.Pdf.dll"

using System;
using System.Linq;
using System.Windows.Forms;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using System.Diagnostics;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

using System.IO;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using System.Reflection;
using System.Collections.Generic;


public static class CitaviMacro
{
	public static void Main()
	{
		
	    // 1. 获取主窗体和项目
		Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
	   // 获取PDF预览控件
		object pdfViewControlObject = null;
		if (Program.ActiveProjectShell.PrimaryMainForm != null)
		{
		    if (Program.ActiveProjectShell.PrimaryMainForm.PreviewControl != null)
		    {
		        PropertyInfo pdfViewControlProperty = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl.GetType().GetProperty("PdfViewControl", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
		        if (pdfViewControlProperty != null)
		        {
		            pdfViewControlObject = pdfViewControlProperty.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl);
		        }
		    }
		}		


		// 步骤 2: 获取 Tool 对象
		Type pdfViewControlType = pdfViewControlObject.GetType();
		PropertyInfo toolProperty = pdfViewControlType.GetProperty("Tool");


		object toolObject = toolProperty.GetValue(pdfViewControlObject);
		if (toolObject == null)
		{
			MessageBox.Show("无法获取Tool对象。");
			return;
		}


	
		//MessageBox.Show("成功：获取到 Tool 对象。\n类型是: " + toolObject.GetType().FullName);

		// 步骤 3: 获取 Tool 的 SelectedAdornmentContainers 字段 (这是核心！)
		Type toolType = toolObject.GetType();
		FieldInfo selectedContainersField = toolType.GetField("SelectedAdornmentContainers", BindingFlags.Instance | BindingFlags.NonPublic);

		object selectedContainersObject = selectedContainersField.GetValue(toolObject);

		if (selectedContainersObject == null)
		{
		    //MessageBox.Show("提示：SelectedAdornmentContainers 字段是 null。");
		    return;
		}
		// 将结果转换为 IEnumerable 以便遍历
		var selectedContainers = selectedContainersObject as System.Collections.IEnumerable;
		// 步骤 4: 遍历每个 AdornmentCanvas，获取其 Annotation 属性
		var annotations = new System.Collections.Generic.List<object>();
		foreach (var container in selectedContainers)
		{
		    if (container == null) continue;

		    Type containerType = container.GetType();
		    PropertyInfo annotationProperty = containerType.GetProperty("Annotation");
		    if (annotationProperty == null) continue;

		    object annotation = annotationProperty.GetValue(container);
		    if (annotation != null)
		    {
		        annotations.Add(annotation);
		    }
		}
		// 步骤 5: 检查结果
		if (annotations.Count == 0)
		{
		    //MessageBox.Show("提示：SelectedAdornmentContainers 集合不为空，但没有找到任何 Annotation 对象。\n请确保您点击的是一个高亮注释，而不是仅仅选中了文本。");
		    return;
		}
		// --- 步骤 3: 从第一个选中的注释获取其Quad区域 ---
		object firstAnnotationObject = annotations.First();
		Annotation selectedAnnotation = firstAnnotationObject as Annotation;

		if (selectedAnnotation == null || selectedAnnotation.Quads == null)
		{
			MessageBox.Show("无法从选中的对象中获取区域信息。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
			return;
		}

		var validQuads = new List<object>();
		foreach (var quad in selectedAnnotation.Quads)
		{
			// 过滤掉容器类型的Quad，只保留实际的文本区域
			PropertyInfo isContainerProp = quad.GetType().GetProperty("IsContainer");
			if (isContainerProp != null && (bool)isContainerProp.GetValue(quad) == false)
			{
				validQuads.Add(quad);
			}
		}

		if (validQuads.Count == 0)
		{
			MessageBox.Show("注释中没有有效的区域可供创建选择。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
			return;
		}

		// --- 步骤 4: 创建新的高亮选择并添加到视图 ---
		try
		{
			// 调用 CreateSelection 方法创建Selection对象
			MethodInfo createSelectionMethod = pdfViewControlType.GetMethod("CreateSelection");
			if (createSelectionMethod == null)
			{
				MessageBox.Show("无法找到创建选择的方法。", "内部错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			object newSelection = createSelectionMethod.Invoke(pdfViewControlObject, new object[] { validQuads });

			if (newSelection != null)
			{
				// 获取视图的Selections集合并将新创建的Selection添加进去
				PropertyInfo selectionsProperty = pdfViewControlType.GetProperty("Selections");
				object selectionsCollection = selectionsProperty.GetValue(pdfViewControlObject);
				
				if (selectionsCollection != null)
				{
					MethodInfo addMethod = selectionsCollection.GetType().GetMethod("Add");
					if (addMethod != null)
					{
						// --- 修改点：不使用 ?. 运算符 ---
						addMethod.Invoke(selectionsCollection, new object[] { newSelection });
						
						MessageBox.Show("成功！已根据选中的注释创建了一个新的高亮选择。", "操作成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("创建选择时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
		
		
	}

}