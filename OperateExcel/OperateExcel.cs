using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace OperateExcel
{
    /// <summary>
    /// エクセル操作ユーティリティ
    /// </summary>
    public class OperateExcel : IDisposable
    {
        // Excel操作用オブジェクト
        protected Application xlApp = null;
        protected Workbooks xlBooks = null;
        protected Workbook xlBook = null;
        protected Sheets xlSheets = null;
        protected Worksheet xlSheet = null;

        /// <summary>
        /// エクセルを開く
        /// </summary>
        /// <param name="excelPath">エクセルファイルのパス（指定しない場合は新規)</param>
        public void Open(string excelPath = null)
        {
            xlApp = new Application();
            xlApp.DisplayAlerts = false;
            xlBooks = xlApp.Workbooks;
            if (excelPath == null)
                xlBook = xlBooks.Add();
            else
                xlBook = xlBooks.Open(CalcPath(excelPath));
            xlSheets = xlBook.Sheets;
            xlSheet = xlSheets[1];
        }

        /// <summary>
        /// 閉じる
        /// </summary>
        public void Close()
        {
            try
            {
                // xlSheet解放
                if (xlSheet != null)
                {
                    Marshal.ReleaseComObject(xlSheet);
                    xlSheet = null;
                }
         
                // xlSheets解放
                if (xlSheets != null)
                {
                    Marshal.ReleaseComObject(xlSheets);
                    xlSheets = null;
                }
         
                // xlBook解放
                if (xlBook != null)
                {
                    try
                    {
                        xlBook.Close();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(xlBook);
                        xlBook = null;
                    }
                }
         
                // xlBooks解放
                if (xlBooks != null)
                {
                    Marshal.ReleaseComObject(xlBooks);
                    xlBooks = null;
                }
         
                // xlApp解放
                if (xlApp != null)
                {
                    try
                    {
                        GC.Collect();
                        // アラートを戻して終了
                        xlApp.DisplayAlerts = true;
                        xlApp.Quit();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(xlApp);
                        xlApp = null;
                        GC.Collect();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// 保存する
        /// </summary>
        /// <param name="savePath">名前を付けて保存する場合は、パスを指定する</param>
        public void Save(string savePath = null)
        {
            if(savePath == null)
                xlBook.Save();
            else
                xlBook.SaveAs(CalcPath(savePath));
        }

        /// <summary>
        /// シートを選択する
        /// </summary>
        /// <param name="sheetName"></param>
        public void SelectSheet(string sheetName)
        {
            var sh = FindSheet(sheetName);
            if(sh == null)
                throw new ArgumentException("ブック内に指定シート名は見つかりません。");

            Marshal.ReleaseComObject(xlSheet);
            xlSheet = sh;
        }

        public enum SheetPosition
        {
            First,
            Last,
            BeforeCurrentSheet,
            AfterCurrentSheet,
        }
        /// <summary>
        /// 末尾のシートを指定移動する
        /// </summary>
        /// <param name="sheetName"></param>
        private void MoveLastSheet(SheetPosition position)
        {
            Worksheet moveSheet = null;
            Worksheet insertTo = null;
            try
            {
                moveSheet = xlSheets[xlSheets.Count];
                switch(position)
                {
                    case SheetPosition.First:
                        insertTo = xlSheets[1];
                        moveSheet.Move(Before: insertTo);
                        break;
                    case SheetPosition.Last:
                        //insertTo = xlSheets[xlSheets.Count];
                        //moveSheet.Move(After: insertTo);
                        break;
                    case SheetPosition.BeforeCurrentSheet:
                        moveSheet.Move(Before: xlSheet);
                        break;
                    case SheetPosition.AfterCurrentSheet:
                        moveSheet.Move(After: xlSheet);
                        break;
                }
            }finally
            {
                if(moveSheet != null)
                    Marshal.ReleaseComObject(moveSheet);
                    moveSheet = null;

                if(insertTo != null)
                    Marshal.ReleaseComObject(insertTo);
                    insertTo = null;
            }
        }
        /// <summary>
        /// シートを作成する
        /// ※カレントシートは変更しない
        /// </summary>
        /// <param name="sheetName"></param>
        public void CreateSheet(string sheetName, SheetPosition position)
        {
            Worksheet lastSheet = null;
            Worksheet newSheet = null;
            try
            {
                lastSheet = xlSheets[xlSheets.Count];
                newSheet = xlSheets.Add(After: lastSheet);
                newSheet.Name = sheetName;
                MoveLastSheet(position);

            }finally
            {
                if(newSheet != null)
                    Marshal.ReleaseComObject(newSheet);
                    newSheet = null;

                if(lastSheet != null)
                    Marshal.ReleaseComObject(lastSheet);
                    lastSheet = null;
            }
        }
        /// <summary>
        /// シートをコピーする
        /// ※カレントシートは変更しない
        /// </summary>
        /// <param name="sheetName"></param>
        public void CopySheet(string toSheetName, string fromSheetName, SheetPosition position)
        {
            Worksheet fromSheet = null;
            Worksheet lastSheet = null;
            Worksheet newSheet = null;
            try
            {
                fromSheet = FindSheet(fromSheetName);
                if(fromSheet == null)
                    throw new ArgumentException("コピーしようとしているシート[" + fromSheetName + "]が見つかりません。");

                //末尾にコピーしてから移動する(Copyメソッドの戻り値がないため、作成したシートインスタンスを取得できない)
                lastSheet = xlSheets[xlSheets.Count];
                fromSheet.Copy(After: lastSheet);
                newSheet = xlSheets[xlSheets.Count];
                newSheet.Name = toSheetName;
                MoveLastSheet(position);

            }finally
            {
                if(newSheet != null)
                    Marshal.ReleaseComObject(newSheet);
                    newSheet = null;

                if(lastSheet != null)
                    Marshal.ReleaseComObject(lastSheet);
                    lastSheet = null;

                if(newSheet != null)
                    Marshal.ReleaseComObject(newSheet);
                    newSheet = null;
            }
        }

        /// <summary>
        /// シート名の一覧を取得する(シートと同じ順番)
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetNames()
        {
            var names = new List<string>();
            //foreach (Worksheet sh in xlSheets)
            for(int i=1; i<=xlSheets.Count; i++)
            {
                var sh = xlSheets[i];
                names.Add(sh.Name);
                Marshal.ReleaseComObject(sh);
            }
            return names;
        }

        /// <summary>
        /// 指定シート名のシートを返す内部関数
        /// ※呼び出し側で、シートインスタンスのReleaseComObject()が必要
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private Worksheet FindSheet(string sheetName)
        {
            foreach (Worksheet sh in xlSheets)
            {
                if (sheetName == sh.Name)
                {
                    return sh;
                }
                Marshal.ReleaseComObject(sh);
            }
            return null;
        }

        /// <summary>
        /// 値を書き込む
        /// </summary>
        /// <param name="row">行番号(0開始)</param>
        /// <param name="column">列番号(0開始)</param>
        /// <param name="values">書き込む値(複数指定した場合は、右セルに順に書き込む)</param>
        public void Write(Cell startCell, params object[] values)
        {
            Range xlRange = null;
            Range xlStartRange = null;
            Range xlEndRange = null;

            try
            {
                xlStartRange = xlSheet.Range[startCell.A1Address];
                xlEndRange = xlStartRange.Offset[0, values.Count() - 1];
                xlRange = xlSheet.Range[xlStartRange, xlEndRange];
                xlRange.Value = values;
            }
            finally
            {
                if (xlRange != null)
                {
                    Marshal.ReleaseComObject(xlRange);
                    xlRange = null;
                }
                if (xlStartRange != null)
                {
                    Marshal.ReleaseComObject(xlStartRange);
                    xlStartRange = null;
                }
                if (xlEndRange != null)
                {
                    Marshal.ReleaseComObject(xlEndRange);
                    xlEndRange = null;
                }
            }
        }

        /// <summary>
        /// 範囲から値を検索する.
        /// </summary>
        /// <param name="range">検索範囲</param>
        /// <param name="targetStr">検索文字列</param>
        /// <param name="foundRow">見つかった行番号</param>
        /// <param name="FoundColumn">見つかった列番号</param>
        /// <param name="isPartialMatch">部分一致検索かどうか</param>
        /// <returns></returns>
        public bool Find(CellsRange searchRange, string keyword, out int foundRow, out int foundColumn, bool isPartialMatch = false)
        {
            Range xlSearchRange = null;
            Range xlFoundRange = null;
            try
            {
                xlSearchRange = xlSheet.Range[searchRange.A1Address];
                XlLookAt xlLookAt = isPartialMatch ? XlLookAt.xlPart : XlLookAt.xlWhole;
                xlFoundRange = xlSearchRange.Find(What:keyword, LookAt:xlLookAt);
                if (xlFoundRange == null)
                {
                    foundRow = 0;
                    foundColumn = 0;
                    return false;
                }

                foundRow = xlFoundRange.Row;
                foundColumn = xlFoundRange.Column;
                return true;

            }
            finally
            {
                if (xlSearchRange != null)
                {
                    Marshal.ReleaseComObject(xlSearchRange);
                    xlSearchRange = null;
                }
                if (xlFoundRange != null)
                {
                    Marshal.ReleaseComObject(xlFoundRange);
                    xlFoundRange = null;
                }
            }
        }
        /// <summary>
        /// 値を読み取る
        /// </summary>
        /// <typeparam name="T">読み取ったときの型</typeparam>
        /// <param name="row">行番号(0開始)</param>
        /// <param name="column">列番号(0開始)</param>
        /// <returns></returns>
        public T Read<T>(Cell cell)
        {
            Range xlRange = null;

            try
            {
                xlRange = xlSheet.Range[cell.A1Address];
                try
                {
                    var converter = TypeDescriptor.GetConverter(typeof(T));
                    if(converter != null)
                    {
                        return (T)converter.ConvertFromString(xlRange.Value2.ToString());
                    }
                }
                catch { }

                return default(T);
            }
            finally
            {
                if (xlRange != null)
                {
                    Marshal.ReleaseComObject(xlRange);
                    xlRange = null;
                }
            }
        }

        /// <summary>
        /// 指定セルに取り消し線が設定されているかを調べる
        /// </summary>
        /// <typeparam name="T">読み取ったときの型</typeparam>
        /// <param name="row">行番号(0開始)</param>
        /// <param name="column">列番号(0開始)</param>
        /// <returns></returns>
        public bool IsStrikethrough(CellsRange range)
        {
            Range xlRange = null;

            try
            {
                xlRange = xlSheet.Range[range.A1Address];
                bool? result = xlRange.Font.Strikethrough as bool?;
                return result ?? false;
            }
            finally
            {
                if (xlRange != null)
                {
                    Marshal.ReleaseComObject(xlRange);
                    xlRange = null;
                }
            }
        }
        public void DeleteSheet(string sheetName)
        {
            Worksheet xlDelSheet = null;
            try
            {
                xlDelSheet = FindSheet(sheetName);
                if (xlSheet.Name == xlDelSheet.Name)
                    throw new ArgumentException("選択シートは削除できません。");

                xlDelSheet.Delete();
            }
            finally
            {
                if (xlDelSheet != null)
                {
                    Marshal.ReleaseComObject(xlDelSheet);
                    xlDelSheet = null;
                }
            }
        }

        /// <summary>
        /// 行をコピー＆挿入する
        /// </summary>
        /// <param name="toStartRow">選択シートの挿入先の行数(0開始)</param>
        /// <param name="fromStartRow">コピー元の開始行(0開始)</param>
        /// <param name="fromCount">コピー元の終了行(0開始)</param>
        /// <param name="fromSheetName">コピー元のシート名(省略した場合は、選択シート)</param>
        public void CopyAndInsertRow(int toStartRow, int fromStartRow, int fromCount, string fromSheetName = null)
        {
            Range xlFromRange = null;
            Range xlToRange = null;
            Worksheet xlFromSheet = null;
            try
            {
                string fromRangeStr = (fromStartRow + 1).ToString() + ":" + (fromStartRow + fromCount).ToString();
                if(fromSheetName != null)
                {
                    xlFromSheet = FindSheet(fromSheetName);
                    xlFromRange = xlFromSheet.Range[fromRangeStr];
                    xlFromRange.Copy();
                }
                else
                {
                    xlFromRange = xlSheet.Range[fromRangeStr];
                    xlFromRange.Copy();
                }
                xlToRange = xlSheet.Rows[toStartRow + 1];
                xlToRange.Insert();
            }
            finally
            {
                if (xlFromRange != null)
                {
                    Marshal.ReleaseComObject(xlFromRange);
                    xlFromRange = null;
                }
                if (xlToRange != null)
                {
                    Marshal.ReleaseComObject(xlToRange);
                    xlFromRange = null;
                }
                if (xlFromSheet != null)
                {
                    Marshal.ReleaseComObject(xlFromSheet);
                    xlFromSheet = null;
                }
            }
        }
        /// <summary>
        /// パスが相対パスの場合は、実行exeのパスから補完する
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private string CalcPath(string path)
        {
            if (System.IO.Path.IsPathRooted(path))
                return path;
            return System.IO.Path.Combine(new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Directory.FullName, path);
        }

        #region IDisposable Support
        private bool disposedValue = false; // 重複する呼び出しを検出するには

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: マネージ状態を破棄します (マネージ オブジェクト)。
                }

                // TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下のファイナライザーをオーバーライドします。
                // TODO: 大きなフィールドを null に設定します。
                Close();

                disposedValue = true;
            }
        }

        // TODO: 上の Dispose(bool disposing) にアンマネージ リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします。
         ~OperateExcel() {
           // このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
           Dispose(false);
         }

        // このコードは、破棄可能なパターンを正しく実装できるように追加されました。
        public void Dispose()
        {
            // このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
            Dispose(true);
            // TODO: 上のファイナライザーがオーバーライドされる場合は、次の行のコメントを解除してください。
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
