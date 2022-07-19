using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XmlFileImportBase.Models;

namespace XmlFileImportBase.Helpers
{
    public static class TaskHelper
    {
        public static void RunTaskWriteTxt(channel ch, string path)
        {
            try
            {
                Task.Run(() => WriteHelper.WriteTxtAsync(ch.Items, path));
            }
            catch(Exception e)
            {
                throw e;
            }
        }
        public static void RunTaskWriteWord(channel ch, string path)
        {
            try
            {
                Task.Run(() => WriteHelper.WriteWordAsync(ch.Items, path));
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
