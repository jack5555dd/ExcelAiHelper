using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Extensibility;

namespace MsExcelAddin
{
    [ComVisible(true)]
    [Guid("1294510F-6F75-4C85-B39D-8AB58C99744A")]
    [ProgId("CursorExcelAddin.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private Excel.Application excelApp;
        private object addInInstance;
        private static readonly string LogFilePath = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), 
            "CursorExcelAddin.log");

        #region 构造函数和日志方法

        public Connect()
        {
            WriteLog("Connect 构造函数被调用");
        }
        
        private void WriteLog(string message)
        {
            try
            {
                File.AppendAllText(LogFilePath, DateTime.Now.ToString() + ": " + message + "\r\n");
            }
            catch
            {
                // 忽略日志写入错误
            }
        }

        #endregion

        #region IDTExtensibility2 实现

        public void OnConnection(object Application, Extensibility.ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                WriteLog("OnConnection 方法被调用，连接模式: " + ConnectMode);
                excelApp = Application as Excel.Application;
                addInInstance = AddInInst;
                
                if (excelApp != null)
                {
                    WriteLog("成功连接到 Excel 应用程序, 版本: " + excelApp.Version);
                }
                else
                {
                    WriteLog("警告：Excel 应用程序对象为空");
                }
            }
            catch (Exception ex)
            {
                WriteLog("OnConnection 发生错误: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("加载项连接错误：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode RemoveMode, ref Array custom)
        {
            WriteLog("OnDisconnection 方法被调用，断开模式: " + RemoveMode);
            
            try
            {
                // 清理资源
                excelApp = null;
                addInInstance = null;
            }
            catch (Exception ex)
            {
                WriteLog("OnDisconnection 发生错误: " + ex.Message);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            WriteLog("OnAddInsUpdate 方法被调用");
        }

        public void OnStartupComplete(ref Array custom)
        {
            WriteLog("OnStartupComplete 方法被调用");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            WriteLog("OnBeginShutdown 方法被调用");
        }

        #endregion

        #region IRibbonExtensibility 实现

        public string GetCustomUI(string ribbonID)
        {
            WriteLog("GetCustomUI 方法被调用，ribbonID: " + ribbonID);
            
            // 返回自定义 Ribbon XML
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                <ribbon>
                    <tabs>
                        <tab id='customTab' label='Cursor Tools'>
                            <group id='mainGroup' label='Main Features'>
                                <button id='btnShowMessage' 
                                        label='Show Message' 
                                        size='large'
                                        onAction='OnShowMessage'
                                        imageMso='HappyFace'/>
                                <button id='btnFormatting' 
                                        label='Format Cells' 
                                        size='normal'
                                        onAction='OnFormatting'
                                        imageMso='FormattingProperties'/>
                            </group>
                            <group id='helpGroup' label='Help'>
                                <button id='btnHelp' 
                                        label='Help Info' 
                                        size='normal'
                                        onAction='OnHelp'
                                        imageMso='Help'/>
                                <button id='btnAbout' 
                                        label='About' 
                                        size='normal'
                                        onAction='OnAbout'
                                        imageMso='Info'/>
                            </group>
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>";
        }

        #endregion

        #region Ribbon 回调方法

        public void OnShowMessage(IRibbonControl control)
        {
            WriteLog("OnShowMessage 方法被调用，控件ID: " + control.Id);
            MessageBox.Show("Cursor Excel Add-in loaded successfully!\r\nProviding enhanced functionality for Excel.", 
                "Cursor Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnFormatting(IRibbonControl control)
        {
            WriteLog("OnFormatting 方法被调用，控件ID: " + control.Id);
            
            try
            {
                if (excelApp != null && excelApp.Selection != null)
                {
                    Excel.Range selection = excelApp.Selection as Excel.Range;
                    if (selection != null)
                    {
                        // 应用简单格式
                        selection.Font.Bold = true;
                        selection.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                        
                        MessageBox.Show("Formatting applied successfully", "Operation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("格式转换错误: " + ex.Message);
                MessageBox.Show("Error while formatting: " + ex.Message, "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnHelp(IRibbonControl control)
        {
            WriteLog("OnHelp 方法被调用，控件ID: " + control.Id);
            MessageBox.Show("Cursor Excel Add-in Help:\n\n" +
                "- Show Message: Display welcome information\n" +
                "- Format Cells: Apply special formatting to selected cells\n\n" +
                "For more help, contact administrator.", 
                "Help Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnAbout(IRibbonControl control)
        {
            WriteLog("OnAbout 方法被调用，控件ID: " + control.Id);
            MessageBox.Show("Cursor Excel Add-in\n" +
                "Version: 1.0.0\n" +
                "Build Date: " + Assembly.GetExecutingAssembly().GetName().Version + "\n\n" +
                "This add-in provides enhanced functionality for Microsoft Excel.", 
                "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion
    }
}
