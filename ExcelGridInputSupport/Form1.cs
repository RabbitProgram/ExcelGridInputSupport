using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelGridInputSupport
{
    public partial class Form1 : Form
    {
        //ウィンドウ情報取得用
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        //IME操作用（全角・半角切り替え）
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int FindWindow(string className, string windowName);
        [System.Runtime.InteropServices.DllImport("Imm32.dll")]
        private static extern int ImmGetDefaultIMEWnd(int hwnd);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int SendMessage(int hWnd, int wMsg, int wParam, byte[] lParam);
        private const int WM_IME_CONTROL = 0x283;
        private const int IMC_SETOPENSTATUS = 0x6;


        int processid_before = -1; //直前にアクティブだったウィンドウのプロセスID
        IntPtr processhandle_before;
        string text_before = "";

        //変換用
        string[] kana_before = { "ゔ", "が", "ぎ", "ぐ", "げ", "ご", "ざ", "じ", "ず", "ぜ", "ぞ", "だ", "ぢ", "づ", "で", "ど", "ば", "び", "ぶ", "べ", "ぼ"
                ,"ヴ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ"
        ,"ぱ","ぴ","ぷ","ぺ","ぽ","パ","ピ","プ","ペ","ポ"};
        string[] kana_after = { "う゛", "か゛", "き゛", "く゛", "け゛", "こ゛", "さ゛", "し゛", "す゛", "せ゛", "そ゛", "た゛", "ち゛", "つ゛", "て゛", "と゛", "は゛", "ひ゛", "ふ゛", "へ゛", "ほ゛"
                ,"ウ゛", "カ゛", "キ゛", "ク゛", "ケ゛", "コ゛", "サ゛", "シ゛", "ス゛", "セ゛", "ソ゛", "タ゛", "チ゛", "ツ゛", "テ゛", "ト゛", "ハ゛", "ヒ゛", "フ゛", "ヘ゛", "ホ゛"
        ,"は゜","ひ゜","ふ゜","へ゜","ほ゜","ハ゜","ヒ゜","フ゜","ヘ゜","ホ゜"};


        public Form1()
        {
            InitializeComponent();
        }

        //左端から入力
        private async void button1_ClickAsync(object sender, EventArgs e)
        {
            SendStringToCharAsync(1);
        }

        //右端から入力
        private void button3_Click(object sender, EventArgs e)
        {
            SendStringToCharAsync(2);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) && (Control.ModifierKeys == Keys.Shift))
            {
                //Shift+Enterキーが押された場合→右端から入力
                button3.PerformClick();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                //Enterキーが押された場合→左端から入力
                button1.PerformClick();
            }
        }

        //アクティブウィンドウの変更を検知
        private void timer1_Tick(object sender, EventArgs e)
        {
            IntPtr hWnd = GetForegroundWindow();
            int id;
            GetWindowThreadProcessId(hWnd, out id);
            string title = Process.GetProcessById(id).MainWindowTitle;

            if (title.Length != 0 && title != this.Text)
            {
                //このソフトのウィンドウじゃない場合は記録
                processid_before = id;
                processhandle_before = Process.GetProcessById(id).MainWindowHandle;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            ActiveControl = textBox1;  //textBox1にフォーカスをセット
        }

        //操作を元に戻す
        private async void button2_ClickAsync(object sender, EventArgs e)
        {
            if (processid_before == -1)
            {
                //他のソフトが一度もアクティブになっていない

            }
            else
            {
                this.Text = "Excel方眼紙入力支援ツール - 元に戻しています...";
                textBox1.Enabled = false;
                spacecount.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;

                //アクティブ化
                Microsoft.VisualBasic.Interaction.AppActivate(processid_before);

                await Task.Delay(200);

                //Ctrl+Zを押す（入力した文字数と、Deleteキーの分）
                for (int i = 1; i <= (text_before.Length) * 2 - 1; i++)
                {
                    SendKeys.Send("^z");
                    await Task.Delay(100);
                }

                textBox1.Text = text_before;
                textBox1.Enabled = true;
                spacecount.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = false;  //元に戻すボタン
                button3.Enabled = true;
                button4.Enabled = true;

                //このソフトをアクティブ化
                Microsoft.VisualBasic.Interaction.AppActivate(this.Text);
                textBox1.Focus();

                this.Text = "Excel方眼紙入力支援ツール";
            }
        }

        //1文字ずつキー送信する（mode→１：左から入力、２：右から入力）
        async void SendStringToCharAsync(int mode)
        {
            int space = int.Parse(spacecount.Value.ToString());  //セルの間隔

            if (processid_before == -1)
            {
                //他のソフトが一度もアクティブになっていない

            }
            else
            {
                this.Text = "Excel方眼紙入力支援ツール - 入力中...";
                textBox1.Enabled = false;
                spacecount.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;

                IMEswitch(false);  //IMEオフ（半角）

                //アクティブ化
                Microsoft.VisualBasic.Interaction.AppActivate(processid_before);

                await Task.Delay(200);

                string moji = textBox1.Text;

                //濁点文字を分解する
                for (int i = 0; i < kana_before.Length; i++)
                {
                    moji = moji.Replace(kana_before[i], kana_after[i]);
                }

                text_before = moji;

                //テキスト要素を列挙するオブジェクトを取得
                TextElementEnumerator charEnum = StringInfo.GetTextElementEnumerator(moji);

                //1文字ずつ解析する
                string[] single = new string[moji.Length];
                StringBuilder analyzedMoji = new StringBuilder();

                for (int i = 0; true; i++)
                {
                    // 次の1文字を取得する
                    if (charEnum.MoveNext() == false)
                    {
                        //取得する文字がない
                        break;
                    }

                    single[i] = charEnum.Current.ToString();
                }

                if (mode == 2)
                {
                    //配列を逆転する
                    Array.Reverse(single);
                }

                for (int a = 0; a < single.Length; a++)
                {
                    //文字を上書き
                    SendKeys.Send("{DELETE}");
                    await Task.Delay(10);

                    //編集モードにする
                    SendKeys.Send("{F2}");
                    await Task.Delay(10);

                    //入力
                    SendKeys.Send(single[a]);
                    await Task.Delay(10);

                    //確定
                    //SendKeys.Send("{ENTER}");
                    //await Task.Delay(50);

                    //カーソルを指定した間隔＋１回分右にずらす
                    for (int i = 1; i <= space + 1; i++)
                    {
                        if (mode == 1)
                        {
                            //右にずらす
                            SendKeys.Send("{TAB}");
                        }
                        else
                        {
                            //左にずらす
                            SendKeys.Send("+{TAB}");
                        }
                        await Task.Delay(10);
                    }
                }

                textBox1.Text = "";
                textBox1.Enabled = true;
                spacecount.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = true;  //元に戻すボタン
                button3.Enabled = true;
                button4.Enabled = true;

                //このソフトをアクティブ化
                Microsoft.VisualBasic.Interaction.AppActivate(this.Text);
                textBox1.Focus();

                this.Text = "Excel方眼紙入力支援ツール";
            }
        }

        //全角・半角切り替え
        private void IMEswitch(bool flg)
        {
            int imeHandle = ImmGetDefaultIMEWnd(processhandle_before.ToInt32());
            int ret;

            if (flg == true)
            {
                ret = SendMessage(imeHandle, WM_IME_CONTROL, IMC_SETOPENSTATUS, new byte[256]);
            }
            else
            {
                ret = SendMessage(imeHandle, WM_IME_CONTROL, IMC_SETOPENSTATUS, null);
            }
        }

        //不透明にする
        private void up_Click(object sender, EventArgs e)
        {
            Opacity = Opacity + 0.1;
        }

        //透明にする
        private void down_Click(object sender, EventArgs e)
        {
            if (Opacity > 0.3)
            {
                Opacity = Opacity - 0.1;
            }
        }

        private async void button4_ClickAsync(object sender, EventArgs e)
        {
            try
            {
                if (processid_before == -1)
                {
                    //他のソフトが一度もアクティブになっていない

                }
                else
                {
                    this.Text = "Excel方眼紙入力支援ツール - 抽出中...";
                    textBox1.Enabled = false;
                    spacecount.Enabled = false;
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button4.Enabled = false;

                    //アクティブ化
                    Microsoft.VisualBasic.Interaction.AppActivate(processid_before);

                    await Task.Delay(200);

                    //Ctrl+Cを押す
                    SendKeys.Send("^c");
                    await Task.Delay(100);

                    //整形
                    string result = Clipboard.GetText().Replace("\t", "").Replace(Environment.NewLine, "");

                    //濁点文字を結合する
                    for (int i = 0; i < kana_before.Length; i++)
                    {
                        result = result.Replace(kana_after[i], kana_before[i]);
                    }

                    //再度クリップボードにコピー
                    Clipboard.SetText(result);

                    textBox1.Text = result;
                    textBox1.Enabled = true;
                    spacecount.Enabled = true;
                    button1.Enabled = true;
                    button2.Enabled = false;  //元に戻すボタン
                    button3.Enabled = true;
                    button4.Enabled = true;

                    //このソフトをアクティブ化
                    Microsoft.VisualBasic.Interaction.AppActivate(this.Text);
                    textBox1.Focus();

                    this.Text = "Excel方眼紙入力支援ツール";
                }
            }
            catch
            {
                textBox1.Text = "抽出時にエラーが発生しました";
                textBox1.Enabled = true;
                spacecount.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = false;  //元に戻すボタン
                button3.Enabled = true;
                button4.Enabled = true;

                //このソフトをアクティブ化
                Microsoft.VisualBasic.Interaction.AppActivate(this.Text);
                textBox1.Focus();

                this.Text = "Excel方眼紙入力支援ツール";
            }
        }
    }
}
