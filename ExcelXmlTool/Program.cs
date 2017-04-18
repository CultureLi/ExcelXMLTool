using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelXmlTool
{
    class Program
    {
        //存储需要格式化的xml文件的名字
        private static List<string> _changeXmlFileNameList = new List<string>();
        //存在于配置文件夹下的xml文件
        private static List<string> _changeXmlFileNameLiveList = new List<string>();
        //当前的工作路径
        private static string _nowWorkPath;
        //格式化工具类
        private static FormatTool _formatTool = new FormatTool();
        //标记是否有格式化
        public static Boolean haveFormat;

        //初始化工作
        private static void workInit()
        {
            _nowWorkPath = System.Environment.CurrentDirectory;
        }

        //有错误信息时显示提示
        private static void showError()
        {
            //退出
            Environment.Exit(0);
        }

        //格式化Xml文件
        private static void formatXmlFile()
        {
            haveFormat = false;
            for (int i = 0; i < _changeXmlFileNameLiveList.Count; i++)
            {
                try
                {
                    File.Move(_changeXmlFileNameLiveList[i], _changeXmlFileNameLiveList[i]);
                }
                catch
                {
                    Console.WriteLine(_changeXmlFileNameList[i] + " error:请关闭xml文件后重试");
                    showError();
                }
                _formatTool.formatXmlFile(_changeXmlFileNameLiveList[i]);
            }
        }

        //检查文件是否存在
        private static Boolean checkFile()
        {
            for (int i = 0; i < _changeXmlFileNameList.Count; i++)
            {
                if (!File.Exists(_nowWorkPath + @"/" + _changeXmlFileNameList[i]))
                {
                    Console.WriteLine(_changeXmlFileNameList[i] + " error:文件不存在");
                }
                else
                {
                    _changeXmlFileNameLiveList.Add(_nowWorkPath + @"/" + _changeXmlFileNameList[i]);
                }
            }
            return true;
        }

        static void Main(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].IndexOf(".xml") < 0)
                {
                    Console.WriteLine("error:文件不存在");
                    showError();
                    return;
                }
                _changeXmlFileNameList.Add(args[i]);
            }
            //初始化
            workInit();
            //检查文件是否存在
            checkFile();
            //格式化文件
            formatXmlFile();
            if (haveFormat)
            {
                //Console.WriteLine("error:first commit");
                haveFormat = false;
            }
            Console.WriteLine("success");
        }
    }
}
